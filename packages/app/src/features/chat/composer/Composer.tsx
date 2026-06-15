import {
  countDomNodes,
  measurePerfAsync,
  recordPerfMetric,
  updatePerfSnapshot,
} from "@better-teams/app/platform/perf";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { teamsThreadService } from "@better-teams/app/services/teams/thread";
import { canonAvatarMri } from "@better-teams/core/teams/profile/avatars";
import { cn } from "@better-teams/ui/utils";
import { useMutation, useQueryClient } from "@tanstack/react-query";
import {
  ArrowUp,
  AtSign,
  Bold,
  Indent,
  Italic,
  Link2,
  List,
  ListOrdered,
  Outdent,
  Paperclip,
  Strikethrough,
} from "lucide-react";
import {
  startTransition,
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";

type LinkDraft = {
  text: string;
  url: string;
};

type PendingAttachment = {
  id: string;
  file: File;
};

function htmlToPlainText(html: string): string {
  if (!html) return "";
  const tmp = document.createElement("div");
  tmp.innerHTML = html;
  return (tmp.textContent ?? "").replace(/\u00a0/g, " ").trim();
}

function normalizeComposerHtml(html: string): string {
  const trimmed = html.trim();
  if (!trimmed || trimmed === "<br>") return "";
  return trimmed
    .replace(/<(div|p)><br><\/(div|p)>/gi, "")
    .replace(/^\s+|\s+$/g, "");
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeAttribute(value: string): string {
  return escapeHtml(value).replaceAll("'", "&#39;");
}

function getEditorSelection(editor: HTMLDivElement | null): Selection | null {
  if (!editor) return null;
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  if (
    !editor.contains(range.startContainer) ||
    !editor.contains(range.endContainer)
  ) {
    return null;
  }
  return selection;
}

function getCaretOffset(editor: HTMLDivElement): number {
  const selection = getEditorSelection(editor);
  if (!selection) return 0;
  const range = selection.getRangeAt(0).cloneRange();
  range.selectNodeContents(editor);
  range.setEnd(selection.anchorNode ?? editor, selection.anchorOffset);
  return range.toString().length;
}

function locateDomPosition(
  root: Node,
  targetOffset: number,
): { node: Node; offset: number } {
  let remaining = targetOffset;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let current = walker.nextNode();
  while (current) {
    const length = current.textContent?.length ?? 0;
    if (remaining <= length) {
      return { node: current, offset: remaining };
    }
    remaining -= length;
    current = walker.nextNode();
  }
  return { node: root, offset: root.childNodes.length };
}

function setSelectionByOffsets(
  editor: HTMLDivElement,
  start: number,
  end: number,
): Range | null {
  const selection = window.getSelection();
  if (!selection) return null;
  const range = document.createRange();
  const startPos = locateDomPosition(editor, Math.max(0, start));
  const endPos = locateDomPosition(editor, Math.max(0, end));
  range.setStart(startPos.node, startPos.offset);
  range.setEnd(endPos.node, endPos.offset);
  selection.removeAllRanges();
  selection.addRange(range);
  return range;
}

function moveCaretToEnd(editor: HTMLDivElement): void {
  const selection = window.getSelection();
  if (!selection) return;
  const range = document.createRange();
  range.selectNodeContents(editor);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}

function ensureEditorSelection(editor: HTMLDivElement): void {
  editor.focus();
  if (getEditorSelection(editor)) return;
  moveCaretToEnd(editor);
}

export type ComposerMentionCandidate = {
  mri: string;
  displayName: string;
  email?: string;
  avatarSrc?: string;
};

type ActiveMention = {
  query: string;
  start: number;
  end: number;
};

function execRichTextCommand(
  command: string,
  value?: string,
  fallback?: () => void,
): void {
  const execCommand = document.execCommand as
    | ((commandId: string, showUI?: boolean, value?: string) => boolean)
    | undefined;
  if (typeof execCommand === "function") {
    const applied = execCommand.call(document, command, false, value);
    if (applied) return;
  }
  fallback?.();
}

function ToolbarIconBtn({
  icon: Icon,
  label,
  onClick,
}: {
  icon: React.ComponentType<{ className?: string }>;
  label: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
      className="flex size-8 items-center justify-center rounded-md text-muted-foreground/65 transition-colors hover:bg-accent hover:text-foreground"
      aria-label={label}
    >
      <Icon className="size-4" />
    </button>
  );
}

export function Composer({
  tenantId,
  conversationId,
  conversationTitle,
  conversationMembers,
  composerRef,
  liveSessionReady,
  mentionCandidates,
}: {
  tenantId?: string | null;
  conversationId: string;
  conversationTitle: string;
  conversationMembers: string[];
  composerRef: React.RefObject<HTMLDivElement | null>;
  liveSessionReady: boolean;
  mentionCandidates: ComposerMentionCandidate[];
}) {
  const queryClient = useQueryClient();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const linkTextInputRef = useRef<HTMLInputElement>(null);
  const [draftHtml, setDraftHtml] = useState("");
  const [activeMention, setActiveMention] = useState<ActiveMention | null>(
    null,
  );
  const deferredMentionQuery = useDeferredValue(activeMention?.query ?? "");
  const [highlightedMentionIndex, setHighlightedMentionIndex] = useState(0);
  const [linkDraft, setLinkDraft] = useState<LinkDraft | null>(null);
  const [pendingAttachments, setPendingAttachments] = useState<
    PendingAttachment[]
  >([]);

  const { mutateAsync: postChatMessage, isPending: sending } = useMutation({
    mutationFn: async (input: {
      conversationId: string;
      content: string;
      contentFormat: "html" | "text";
      mentions: Array<Record<string, unknown>>;
      attachments: File[];
      conversationMembers: string[];
    }) =>
      measurePerfAsync(
        "composer.send",
        {
          conversationId: input.conversationId,
          attachmentCount: input.attachments.length,
          mentionCount: input.mentions.length,
          contentLength: input.content.length,
        },
        async () => {
          await teamsThreadService.sendMessage(tenantId, input);
        },
      ),
    onSuccess: async (_data, input) => {
      await Promise.all([
        queryClient.invalidateQueries({
          queryKey: teamsKeys.thread(tenantId, input.conversationId),
        }),
        queryClient.invalidateQueries({
          queryKey: teamsKeys.conversations(tenantId),
        }),
      ]);
    },
  });

  useEffect(() => {
    if (!composerRef.current) return;
    if (composerRef.current.innerHTML !== draftHtml) {
      composerRef.current.innerHTML = draftHtml;
    }
  }, [composerRef, draftHtml]);

  const syncDraftFromDom = useCallback(() => {
    const html = normalizeComposerHtml(composerRef.current?.innerHTML ?? "");
    recordPerfMetric("composer.syncDraft", {
      conversationId,
      contentLength: html.length,
    });
    setDraftHtml(html);
  }, [composerRef, conversationId]);

  const syncMentionState = useCallback(() => {
    const editor = composerRef.current;
    if (!editor) {
      setActiveMention(null);
      return;
    }
    const selection = getEditorSelection(editor);
    if (!selection || !selection.isCollapsed) {
      setActiveMention(null);
      return;
    }
    const caretOffset = getCaretOffset(editor);
    const beforeCaret = editor.innerText.slice(0, caretOffset);
    const match = beforeCaret.match(/(?:^|\s)@([^\s@]*)$/);
    if (!match) {
      setActiveMention(null);
      return;
    }
    const query = match[1] ?? "";
    const start = caretOffset - query.length - 1;
    startTransition(() => {
      setActiveMention({ query, start, end: caretOffset });
      setHighlightedMentionIndex(0);
    });
  }, [composerRef]);

  const collectMentionsFromDom = useCallback(() => {
    const editor = composerRef.current;
    if (!editor) return [] as Array<Record<string, unknown>>;
    const nodes = [...editor.querySelectorAll("at[data-mention-mri]")];
    return nodes.map((node, index) => {
      const mentionId = node.getAttribute("id") || `mention-${index}`;
      const displayName =
        node.getAttribute("data-mention-name") ||
        node.textContent?.replace(/^@/, "") ||
        "Unknown";
      const mri = node.getAttribute("data-mention-mri") || "";
      return {
        id: mentionId,
        mri,
        displayName,
        mentionType: "person",
      };
    });
  }, [composerRef]);

  const insertHtml = useCallback(
    (html: string, plainTextFallback: string) => {
      const editor = composerRef.current;
      if (!editor) return;
      ensureEditorSelection(editor);
      execRichTextCommand("insertHTML", html, () => {
        execRichTextCommand("insertText", plainTextFallback);
      });
      syncDraftFromDom();
      syncMentionState();
    },
    [composerRef, syncDraftFromDom, syncMentionState],
  );

  const runCommand = useCallback(
    (command: string, value?: string, fallback?: () => void) => {
      const editor = composerRef.current;
      if (!editor) return;
      ensureEditorSelection(editor);
      execRichTextCommand(command, value, fallback);
      syncDraftFromDom();
      syncMentionState();
    },
    [composerRef, syncDraftFromDom, syncMentionState],
  );

  const applyMention = useCallback(
    (candidate: ComposerMentionCandidate) => {
      const editor = composerRef.current;
      const mention = activeMention;
      if (!editor || !mention) return;
      editor.focus();
      setSelectionByOffsets(editor, mention.start, mention.end);
      const mentionId = `mention-${Date.now()}`;
      const safeName = escapeHtml(candidate.displayName);
      const html = `<at id="${mentionId}" data-mention-mri="${escapeAttribute(
        canonAvatarMri(candidate.mri),
      )}" data-mention-name="${escapeAttribute(candidate.displayName)}">@${safeName}</at>&nbsp;`;
      execRichTextCommand("insertHTML", html, () => {
        execRichTextCommand("insertText", `@${candidate.displayName} `);
      });
      syncDraftFromDom();
      setActiveMention(null);
    },
    [activeMention, composerRef, syncDraftFromDom],
  );

  const openLinkModal = useCallback(() => {
    const editor = composerRef.current;
    if (!editor) return;
    ensureEditorSelection(editor);
    const selection = getEditorSelection(editor);
    const selectedText = selection?.toString().trim() ?? "";
    setLinkDraft({
      text: selectedText,
      url: "https://",
    });
  }, [composerRef]);

  const applyLink = useCallback(() => {
    if (!linkDraft) return;
    const url = linkDraft.url.trim();
    const label = linkDraft.text.trim() || url;
    if (!url) return;
    insertHtml(
      `<a href="${escapeAttribute(url)}">${escapeHtml(label)}</a>`,
      label,
    );
    setLinkDraft(null);
  }, [insertHtml, linkDraft]);

  const openAttachmentPicker = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const removePendingAttachment = useCallback((id: string) => {
    setPendingAttachments((current) =>
      current.filter((attachment) => attachment.id !== id),
    );
  }, []);

  const handleAttachmentSelection = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const files = [...(event.target.files ?? [])];
      if (files.length === 0) return;
      setPendingAttachments((current) => [
        ...current,
        ...files.map((file, index) => ({
          id: `${Date.now()}-${index}-${file.name}`,
          file,
        })),
      ]);
      event.target.value = "";
    },
    [],
  );

  const handleShortcut = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      const shortcut = event.key.toLowerCase();
      const primary = event.metaKey || event.ctrlKey;
      if (!primary) return false;
      if (shortcut === "b") {
        event.preventDefault();
        runCommand("bold");
        return true;
      }
      if (shortcut === "i") {
        event.preventDefault();
        runCommand("italic");
        return true;
      }
      if (shortcut === "u") {
        event.preventDefault();
        runCommand("insertUnorderedList");
        return true;
      }
      if (shortcut === "7" && event.shiftKey) {
        event.preventDefault();
        runCommand("insertOrderedList");
        return true;
      }
      if (shortcut === "8" && event.shiftKey) {
        event.preventDefault();
        runCommand("insertUnorderedList");
        return true;
      }
      if (shortcut === "k") {
        event.preventDefault();
        openLinkModal();
        return true;
      }
      return false;
    },
    [openLinkModal, runCommand],
  );

  const handleSend = useCallback(async () => {
    if (!liveSessionReady || sending) return;
    const html = normalizeComposerHtml(composerRef.current?.innerHTML ?? "");
    const text = htmlToPlainText(html);
    if (!text && pendingAttachments.length === 0) return;
    const mentions = collectMentionsFromDom();
    const attachments = pendingAttachments.map((attachment) => attachment.file);
    recordPerfMetric("composer.sendRequested", {
      conversationId,
      attachmentCount: attachments.length,
      mentionCount: mentions.length,
      contentLength: (html || text).length,
    });
    setDraftHtml("");
    setPendingAttachments([]);
    if (composerRef.current) composerRef.current.innerHTML = "";
    try {
      await postChatMessage({
        conversationId,
        content: html || text,
        contentFormat: html.includes("<") ? "html" : "text",
        mentions,
        attachments,
        conversationMembers,
      });
      requestAnimationFrame(() => composerRef.current?.focus());
    } catch {
      setDraftHtml(html);
      setPendingAttachments(
        attachments.map((file, index) => ({
          id: `${Date.now()}-${index}-${file.name}`,
          file,
        })),
      );
      if (composerRef.current) composerRef.current.innerHTML = html;
    }
  }, [
    composerRef,
    conversationMembers,
    conversationId,
    liveSessionReady,
    pendingAttachments,
    postChatMessage,
    sending,
    collectMentionsFromDom,
  ]);

  const hasText = useMemo(
    () => htmlToPlainText(draftHtml).length > 0,
    [draftHtml],
  );
  const hasAttachments = pendingAttachments.length > 0;
  const filteredMentionCandidates = useMemo(() => {
    const query = deferredMentionQuery.trim().toLowerCase();
    const unique = new Map<string, ComposerMentionCandidate>();
    for (const candidate of mentionCandidates) {
      const key = canonAvatarMri(candidate.mri);
      if (!unique.has(key)) unique.set(key, candidate);
    }
    const items = [...unique.values()];
    if (!query) return items.slice(0, 6);
    return items
      .filter((candidate) => {
        const haystack =
          `${candidate.displayName} ${candidate.email ?? ""}`.toLowerCase();
        return haystack.includes(query);
      })
      .slice(0, 6);
  }, [deferredMentionQuery, mentionCandidates]);

  const activeHighlightedMentionIndex =
    highlightedMentionIndex < filteredMentionCandidates.length
      ? highlightedMentionIndex
      : 0;

  useEffect(() => {
    if (!linkDraft) return;
    linkTextInputRef.current?.focus();
  }, [linkDraft]);

  useEffect(() => {
    updatePerfSnapshot(`composer:${conversationId}`, {
      draftTextLength: htmlToPlainText(draftHtml).length,
      draftHtmlLength: draftHtml.length,
      attachmentCount: pendingAttachments.length,
      mentionCandidateCount: mentionCandidates.length,
      visibleMentionCount: filteredMentionCandidates.length,
      domNodeCount: countDomNodes(composerRef.current),
      sending: sending ? 1 : 0,
    });
  }, [
    composerRef,
    conversationId,
    draftHtml,
    filteredMentionCandidates.length,
    mentionCandidates.length,
    pendingAttachments.length,
    sending,
  ]);

  return (
    <div className="shrink-0 px-5 pt-1 pb-4">
      <form
        onSubmit={(event) => {
          event.preventDefault();
          void handleSend();
        }}
        className="relative z-30 mx-auto max-w-none overflow-visible rounded-xl border border-border bg-background transition-shadow focus-within:border-muted-foreground/20 focus-within:shadow-sm"
      >
        <div className="relative px-4 pt-3 pb-2">
          <input
            ref={fileInputRef}
            type="file"
            multiple
            className="hidden"
            onChange={handleAttachmentSelection}
          />
          {pendingAttachments.length > 0 ? (
            <div className="mb-3 flex flex-wrap gap-2">
              {pendingAttachments.map((attachment) => (
                <span
                  key={attachment.id}
                  className="inline-flex max-w-full items-center gap-2 rounded-full border border-border bg-accent/40 px-3 py-1 text-[12px] text-foreground"
                >
                  <span className="truncate">{attachment.file.name}</span>
                  <button
                    type="button"
                    onClick={() => removePendingAttachment(attachment.id)}
                    className="text-muted-foreground transition-colors hover:text-foreground"
                    aria-label={`Remove ${attachment.file.name}`}
                    disabled={sending}
                  >
                    ×
                  </button>
                </span>
              ))}
            </div>
          ) : null}
          {/* biome-ignore lint/a11y/useSemanticElements: rich contentEditable composer cannot be a textarea. */}
          <div
            {...({ placeholder: `Message ${conversationTitle}…` } as {
              placeholder: string;
            })}
            ref={composerRef}
            role="textbox"
            aria-label="Message text"
            aria-multiline="true"
            tabIndex={0}
            contentEditable={liveSessionReady}
            suppressContentEditableWarning
            data-placeholder={`Message ${conversationTitle}…`}
            onInput={() => {
              syncDraftFromDom();
              syncMentionState();
            }}
            onClick={syncMentionState}
            onKeyUp={syncMentionState}
            onKeyDown={(event) => {
              if (handleShortcut(event)) return;
              if (activeMention && filteredMentionCandidates.length > 0) {
                if (event.key === "ArrowDown") {
                  event.preventDefault();
                  setHighlightedMentionIndex(
                    (current) =>
                      (current + 1) % filteredMentionCandidates.length,
                  );
                  return;
                }
                if (event.key === "ArrowUp") {
                  event.preventDefault();
                  setHighlightedMentionIndex(
                    (current) =>
                      (current - 1 + filteredMentionCandidates.length) %
                      filteredMentionCandidates.length,
                  );
                  return;
                }
                if (event.key === "Enter" && !event.shiftKey) {
                  event.preventDefault();
                  const candidate =
                    filteredMentionCandidates[activeHighlightedMentionIndex];
                  if (candidate) applyMention(candidate);
                  return;
                }
                if (event.key === "Escape") {
                  event.preventDefault();
                  setActiveMention(null);
                  return;
                }
              }
              if (
                event.key === "Enter" &&
                !event.shiftKey &&
                !event.nativeEvent.isComposing
              ) {
                event.preventDefault();
                void handleSend();
              }
            }}
            className={cn(
              "max-h-32 min-h-[48px] w-full overflow-y-auto whitespace-pre-wrap break-words bg-transparent text-[14px] leading-relaxed text-foreground outline-none empty:before:pointer-events-none empty:before:text-muted-foreground/35 empty:before:content-[attr(data-placeholder)] [&_ol]:my-1 [&_ol]:ml-5 [&_ol]:list-decimal [&_ol]:pl-2 [&_ul]:my-1 [&_ul]:ml-5 [&_ul]:list-disc [&_ul]:pl-2 [&_li]:my-0.5",
              !liveSessionReady && "cursor-not-allowed opacity-60",
            )}
          />
          {activeMention && filteredMentionCandidates.length > 0 ? (
            <div className="absolute inset-x-4 bottom-full z-[80] mb-2 rounded-xl border border-border bg-background p-1 shadow-lg">
              {filteredMentionCandidates.map((candidate, index) => (
                <button
                  key={candidate.mri}
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => applyMention(candidate)}
                  className={cn(
                    "flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left transition-colors",
                    index === activeHighlightedMentionIndex
                      ? "bg-accent text-foreground"
                      : "hover:bg-accent/60",
                  )}
                >
                  <span className="flex size-6 shrink-0 items-center justify-center overflow-hidden rounded-full bg-accent text-[10px] font-semibold text-muted-foreground">
                    {candidate.avatarSrc ? (
                      <img
                        src={candidate.avatarSrc}
                        alt=""
                        className="size-full object-cover"
                      />
                    ) : (
                      candidate.displayName
                        .split(/\s+/)
                        .map((part) => part[0] ?? "")
                        .join("")
                        .slice(0, 2)
                        .toUpperCase()
                    )}
                  </span>
                  <span className="min-w-0 flex-1">
                    <span className="block truncate text-[12px] font-medium leading-tight">
                      {candidate.displayName}
                    </span>
                    {candidate.email ? (
                      <span className="block truncate text-[10px] leading-tight text-muted-foreground">
                        {candidate.email}
                      </span>
                    ) : null}
                  </span>
                </button>
              ))}
            </div>
          ) : null}
        </div>

        <div className="flex items-center justify-between border-t border-border px-1 py-0.5">
          <div className="flex items-center">
            <ToolbarIconBtn
              icon={Paperclip}
              label="Attach"
              onClick={openAttachmentPicker}
            />
            <ToolbarIconBtn
              icon={AtSign}
              label="Mention"
              onClick={() => {
                runCommand("insertText", "@");
              }}
            />
            <ToolbarIconBtn
              icon={Link2}
              label="Add link"
              onClick={openLinkModal}
            />
            <div className="mx-1 h-4 w-px bg-border" />
            <ToolbarIconBtn
              icon={Bold}
              label="Bold"
              onClick={() => runCommand("bold")}
            />
            <ToolbarIconBtn
              icon={Italic}
              label="Italic"
              onClick={() => runCommand("italic")}
            />
            <ToolbarIconBtn
              icon={Strikethrough}
              label="Strikethrough"
              onClick={() => runCommand("strikeThrough")}
            />
            <div className="mx-1 h-4 w-px bg-border" />
            <ToolbarIconBtn
              icon={ListOrdered}
              label="Numbered list"
              onClick={() => runCommand("insertOrderedList")}
            />
            <ToolbarIconBtn
              icon={List}
              label="Bullet list"
              onClick={() => runCommand("insertUnorderedList")}
            />
            <ToolbarIconBtn
              icon={Indent}
              label="Indent"
              onClick={() => runCommand("indent")}
            />
            <ToolbarIconBtn
              icon={Outdent}
              label="Outdent"
              onClick={() => runCommand("outdent")}
            />
          </div>
          <button
            type="submit"
            disabled={
              !liveSessionReady || sending || (!hasText && !hasAttachments)
            }
            aria-label="Send"
            className={cn(
              "mr-1 flex size-8 shrink-0 items-center justify-center rounded-lg transition-all duration-150",
              hasText || hasAttachments
                ? "bg-primary text-primary-foreground shadow-sm hover:bg-primary/90"
                : "bg-muted text-muted-foreground/30",
            )}
          >
            <ArrowUp className="size-4" strokeWidth={2.5} />
          </button>
        </div>
      </form>
      {linkDraft ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/10 px-4 backdrop-blur-[2px]">
          <button
            type="button"
            aria-label="Close link dialog"
            className="absolute inset-0"
            onClick={() => setLinkDraft(null)}
          />
          <div
            role="dialog"
            aria-modal="true"
            aria-label="Insert link"
            className="relative w-full max-w-sm rounded-2xl border border-border bg-background p-4 shadow-lg"
          >
            <h3 className="text-[15px] font-semibold">Insert link</h3>
            <div className="mt-4 space-y-3">
              <label className="block">
                <span className="mb-1 block text-[12px] text-muted-foreground">
                  Link text
                </span>
                <input
                  ref={linkTextInputRef}
                  type="text"
                  value={linkDraft.text}
                  onChange={(event) =>
                    setLinkDraft((current) =>
                      current
                        ? { ...current, text: event.target.value }
                        : current,
                    )
                  }
                  className="w-full rounded-lg border border-border bg-background px-3 py-2 text-[13px] outline-none"
                />
              </label>
              <label className="block">
                <span className="mb-1 block text-[12px] text-muted-foreground">
                  Link URL
                </span>
                <input
                  type="url"
                  value={linkDraft.url}
                  onChange={(event) =>
                    setLinkDraft((current) =>
                      current
                        ? { ...current, url: event.target.value }
                        : current,
                    )
                  }
                  className="w-full rounded-lg border border-border bg-background px-3 py-2 text-[13px] outline-none"
                />
              </label>
            </div>
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setLinkDraft(null)}
                className="rounded-lg border border-border px-3 py-2 text-[13px] transition-colors hover:bg-accent"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={applyLink}
                className="rounded-lg bg-primary px-3 py-2 text-[13px] text-primary-foreground transition-colors hover:bg-primary/90"
              >
                Insert
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
