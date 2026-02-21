/**
 * Word Document Operations for Armenian Name Indexer
 * Handles interaction with Word document via Office.js API
 */

/* global Word */

import { IndexerSettings, IndexResult, NameMatch, ProgressCallback } from "../types";
import { 
  parseArmenianName, 
  normalizeArmenianSurname, 
  parseExceptionsList, 
  isExcluded 
} from "./armenian";

/**
 * Find all Armenian names in the document matching the pattern
 * Uses paragraph-by-paragraph search since Office.js doesn't support regex
 * 
 * @param context - Word request context
 * @param pattern - Regex pattern to match
 * @returns Promise resolving to array of name matches
 */
export async function findArmenianNamesInDocument(
  context: Word.RequestContext,
  pattern: RegExp
): Promise<NameMatch[]> {
  const matches: NameMatch[] = [];
  
  // Get all paragraphs
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  let globalOffset = 0;

  // Search each paragraph
  for (let i = 0; i < paragraphs.items.length; i++) {
    const paragraph = paragraphs.items[i];
    paragraph.load("text");
    await context.sync();

    const text = paragraph.text;
    
    // Reset regex lastIndex for each paragraph
    pattern.lastIndex = 0;
    
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text)) !== null) {
      
      // Get the actual range by searching for the matched text
      // This is more reliable than character offset calculations
      const searchResults = paragraph.search(match[0], {
        matchCase: true,
        matchWholeWord: false
      });
      searchResults.load("items");
      await context.sync();
      
      if (searchResults.items.length > 0) {
        // Use the first search result that matches our position
        const foundRange = searchResults.items[0];
        
        matches.push({
          text: match[0],
          range: foundRange,
          startIndex: globalOffset + match.index,
          length: match[0].length
        });
      }
    }
    
    // Update global offset for next paragraph
    globalOffset += text.length;
  }

  return matches;
}

/**
 * Mark an index entry at the specified range
 * Inserts an XE (index entry) field
 * 
 * @param context - Word request context
 * @param range - Range where to insert the index entry
 * @param entry - Index entry text
 */
export async function markIndexEntry(
  context: Word.RequestContext,
  range: Word.Range,
  entry: string
): Promise<void> {
  // Insert XE field before the matched text
  // XE field format: { XE "entry text" }
  try {
    range.insertField(Word.InsertLocation.before, "XE", `"${entry}"`);
  } catch (error) {
    // Fallback: insert as text if insertField doesn't work for XE
    const fieldCode = `{ XE "${entry}" }`;
    range.insertText(fieldCode, Word.InsertLocation.before);
  }
}

/**
 * Clear all index entries from the document
 * Ported from VBA ClearAllIndexEntries
 * 
 * @param context - Word request context
 * @param onProgress - Optional progress callback
 * @returns Promise resolving to number of entries deleted
 */
export async function clearAllIndexEntries(
  context: Word.RequestContext,
  onProgress?: ProgressCallback
): Promise<number> {
  const fields = context.document.body.fields;
  fields.load("items");
  await context.sync();

  const total = fields.items.length;
  
  if (total === 0) {
    return 0;
  }

  let deleted = 0;

  // Iterate backwards to avoid index shifting issues
  for (let i = total - 1; i >= 0; i--) {
    const field = fields.items[i];
    field.load("code");
    await context.sync();

    // Check if field code contains XE (index entry)
    // Note: FieldType enum may not have fieldIndexEntry, so check by code
    const fieldCode = field.code || "";
    if (fieldCode.trim().toUpperCase().startsWith("XE")) {
      field.delete();
      deleted++;
      
      // Sync every 10 deletions to avoid too many pending operations
      if (deleted % 10 === 0) {
        await context.sync();
        
        if (onProgress) {
          const percent = Math.floor(((total - i) / total) * 100);
          onProgress(percent, `Clearing index entries… ${deleted} removed`);
        }
      }
    }
  }

  // Final sync
  await context.sync();

  if (onProgress) {
    onProgress(100, `Cleared ${deleted} index entries`);
  }

  return deleted;
}

/**
 * Index all Armenian names in the document
 * Main orchestration function ported from VBA AutoIndexArmenianNames
 * 
 * @param context - Word request context
 * @param settings - Indexer settings
 * @param onProgress - Optional progress callback
 * @returns Promise resolving to index result
 */
export async function indexArmenianNames(
  context: Word.RequestContext,
  settings: IndexerSettings,
  onProgress?: ProgressCallback
): Promise<IndexResult> {
  const result: IndexResult = {
    indexed: 0,
    skipped: 0,
    errors: []
  };

  try {
    // Parse exceptions list
    const exceptionsText = settings.exceptions.join("\n");
    const exceptions = parseExceptionsList(exceptionsText);

    if (onProgress) {
      onProgress(0, "Searching for Armenian names…");
    }

    // Create regex pattern from settings
    const pattern = new RegExp(settings.pattern, "g");

    // Find all matching names
    const matches = await findArmenianNamesInDocument(context, pattern);

    if (matches.length === 0) {
      if (onProgress) {
        onProgress(100, "No Armenian names found");
      }
      return result;
    }

    if (onProgress) {
      onProgress(10, `Found ${matches.length} potential names`);
    }

    // Process matches in reverse order (like VBA) to avoid position shifting
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      const fullName = match.text;

      try {
        // Check if any word is in exceptions list
        if (isExcluded(fullName, exceptions)) {
          result.skipped++;
          continue;
        }

        // Parse name into first and last
        const parsed = parseArmenianName(fullName);
        
        if (!parsed.firstName || !parsed.lastName) {
          result.skipped++;
          continue;
        }

        // Normalize surname
        const normalizedSurname = normalizeArmenianSurname(
          parsed.lastName,
          settings.suffixes
        );

        // Build normalized name (firstName + normalized surname, no patronymic)
        const normalizedName = `${parsed.firstName} ${normalizedSurname}`;

        // Mark index entry
        await markIndexEntry(context, match.range, normalizedName);
        result.indexed++;

        // Sync every 5 entries to balance performance and responsiveness
        if (result.indexed % 5 === 0) {
          await context.sync();
          
          if (onProgress) {
            const percent = Math.floor(10 + ((matches.length - i) / matches.length) * 85);
            onProgress(
              percent, 
              `Indexing… ${result.indexed} indexed, ${result.skipped} skipped`
            );
          }
        }
      } catch (error) {
        result.errors.push(`Error processing "${fullName}": ${error.message}`);
        result.skipped++;
      }
    }

    // Final sync
    await context.sync();

    if (onProgress) {
      onProgress(
        100, 
        `Complete: ${result.indexed} indexed, ${result.skipped} skipped`
      );
    }

  } catch (error) {
    result.errors.push(`Indexing error: ${error.message}`);
    throw error;
  }

  return result;
}
