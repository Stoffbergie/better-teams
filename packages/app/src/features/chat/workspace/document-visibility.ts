import { useEffect, useRef, useState, useSyncExternalStore } from "react";

const RESUME_COOLDOWN_MS = 3_000;

function isDocumentVisible(): boolean {
  return (
    typeof document === "undefined" || document.visibilityState === "visible"
  );
}

function subscribeToDocumentVisibility(onStoreChange: () => void): () => void {
  if (typeof document === "undefined") return () => undefined;
  document.addEventListener("visibilitychange", onStoreChange);
  return () => document.removeEventListener("visibilitychange", onStoreChange);
}

export function useDocumentVisibility(): boolean {
  return useSyncExternalStore(
    subscribeToDocumentVisibility,
    isDocumentVisible,
    () => true,
  );
}

export function useResumeCooldown(delayMs = RESUME_COOLDOWN_MS): boolean {
  const documentVisible = useDocumentVisibility();
  const [resumeReady, setResumeReady] = useState(documentVisible);
  const previousVisibleRef = useRef(documentVisible);

  if (previousVisibleRef.current !== documentVisible) {
    previousVisibleRef.current = documentVisible;
    if (!documentVisible && resumeReady) {
      setResumeReady(false);
    }
  }

  useEffect(() => {
    if (!documentVisible) return;
    const timer = window.setTimeout(() => {
      setResumeReady(true);
    }, delayMs);
    return () => window.clearTimeout(timer);
  }, [delayMs, documentVisible]);

  return resumeReady;
}
