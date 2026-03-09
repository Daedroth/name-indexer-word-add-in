/* global document, Office, console, HTMLElement, HTMLButtonElement */

import "../taskpane/taskpane.css";

type ParentMessage =
  | { type: "progress"; percent: number; status: string }
  | { type: "warning"; message: string }
  | { type: "complete"; message: string; severity?: "success" | "info" | "warning" | "error" }
  | { type: "error"; message: string };

type ChildMessage = { type: "cancel" };

function toErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function updateProgress(percent: number, status: string) {
  const bar = document.getElementById("progress-bar") as HTMLElement;
  const text = document.getElementById("status-text") as HTMLElement;

  const clamped = Math.max(0, Math.min(100, Math.floor(percent)));
  bar.style.width = clamped + "%";
  bar.setAttribute("aria-valuenow", clamped.toString());
  text.textContent = status;
}

function showWarning(message: string) {
  const section = document.getElementById("warning-section") as HTMLElement;
  const div = document.getElementById("warning-message") as HTMLElement;
  div.textContent = message;
  section.style.display = "block";
}

function showResult(message: string, severity: "success" | "info" | "warning" | "error") {
  const resultSection = document.getElementById("result-section") as HTMLElement;
  const resultMessage = document.getElementById("result-message") as HTMLElement;
  const closeBtn = document.getElementById("close-btn") as HTMLButtonElement;

  resultMessage.textContent = message;
  resultMessage.className = "message " + severity;
  resultSection.style.display = "block";

  closeBtn.style.display = "inline-block";
}

function setCancelEnabled(enabled: boolean) {
  const cancelBtn = document.getElementById("cancel-btn") as HTMLButtonElement;
  cancelBtn.disabled = !enabled;
}

function sendToParent(msg: ChildMessage) {
  try {
    Office.context.ui.messageParent(JSON.stringify(msg));
  } catch (error) {
    console.error("Failed to message parent:", toErrorMessage(error));
  }
}

Office.onReady(() => {
  const init = () => {
    (document.getElementById("app-body") as HTMLElement).style.display = "block";

    (document.getElementById("cancel-btn") as HTMLButtonElement).addEventListener("click", () => {
      setCancelEnabled(false);
      updateProgress(0, "Cancelling…");
      sendToParent({ type: "cancel" });
    });

    (document.getElementById("close-btn") as HTMLButtonElement).addEventListener("click", () => {
      try {
        window.close();
      } catch {
        // ignore
      }
    });

    updateProgress(0, "Starting…");

    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      try {
        const data = JSON.parse(arg.message) as ParentMessage;

        if (data.type === "progress") {
          updateProgress(data.percent, data.status);
          return;
        }

        if (data.type === "warning") {
          showWarning(data.message);
          return;
        }

        if (data.type === "complete") {
          setCancelEnabled(false);
          updateProgress(100, "Done");
          showResult(data.message, data.severity ?? "success");
          return;
        }

        if (data.type === "error") {
          setCancelEnabled(false);
          showResult(data.message, "error");
          return;
        }
      } catch (error) {
        console.error("Failed to handle parent message:", toErrorMessage(error));
      }
    });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
});
