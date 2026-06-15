import {
  type RefObject,
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import type { DisplayMessage } from "./types";

type UseThreadSearchOptions = {
  viewportRef: RefObject<HTMLDivElement | null>;
  displayMessages: DisplayMessage[];
  searchQuery: string;
  onSearchResultCountChange?: (resultCount: number) => void;
};

export type ThreadSearchController = {
  highlightedMessageId: string | null;
  matchingMessageIds: string[];
  normalizedSearchQuery: string;
  scrollToMessage: (messageId: string) => void;
  submitSearch: (query: string) => void;
};

export function useThreadSearch({
  viewportRef,
  displayMessages,
  searchQuery,
  onSearchResultCountChange,
}: UseThreadSearchOptions): ThreadSearchController {
  const deferredSearchQuery = useDeferredValue(searchQuery);
  const [highlightedMessageId, setHighlightedMessageId] = useState<
    string | null
  >(null);
  const highlightTimerRef = useRef<number | null>(null);
  const searchMatchStateRef = useRef<{ query: string; index: number }>({
    query: "",
    index: -1,
  });

  const normalizedSearchQuery = deferredSearchQuery.trim().toLowerCase();

  const findMatchingMessageIds = useCallback(
    (normalizedQuery: string) => {
      if (!normalizedQuery) return [];
      return displayMessages
        .filter((entry) => entry.searchText.includes(normalizedQuery))
        .map((entry) => entry.message.id);
    },
    [displayMessages],
  );

  const matchingMessageIds = useMemo(
    () => findMatchingMessageIds(normalizedSearchQuery),
    [findMatchingMessageIds, normalizedSearchQuery],
  );

  const scrollToMessage = useCallback(
    (messageId: string) => {
      const viewport = viewportRef.current;
      const target = [
        ...(viewport?.querySelectorAll<HTMLElement>("[data-message-id]") ?? []),
      ].find((node) => node.dataset.messageId === messageId);
      if (!target) return;
      setHighlightedMessageId(messageId);
      if (highlightTimerRef.current != null) {
        window.clearTimeout(highlightTimerRef.current);
      }
      highlightTimerRef.current = window.setTimeout(() => {
        setHighlightedMessageId((current) =>
          current === messageId ? null : current,
        );
        highlightTimerRef.current = null;
      }, 2200);
      requestAnimationFrame(() => {
        target.scrollIntoView({ block: "center" });
      });
    },
    [viewportRef],
  );

  const submitSearch = useCallback(
    (query: string) => {
      const normalizedQuery = query.trim().toLowerCase();
      if (!normalizedQuery) return;
      const immediateMatches = findMatchingMessageIds(normalizedQuery);
      if (immediateMatches.length === 0) return;
      const nextIndex =
        searchMatchStateRef.current.query === normalizedQuery
          ? (searchMatchStateRef.current.index + 1) % immediateMatches.length
          : 0;
      searchMatchStateRef.current = {
        query: normalizedQuery,
        index: nextIndex,
      };
      const messageId = immediateMatches[nextIndex];
      if (messageId) {
        scrollToMessage(messageId);
      }
    },
    [findMatchingMessageIds, scrollToMessage],
  );

  useEffect(() => {
    onSearchResultCountChange?.(
      normalizedSearchQuery ? matchingMessageIds.length : 0,
    );
  }, [
    matchingMessageIds.length,
    normalizedSearchQuery,
    onSearchResultCountChange,
  ]);

  useEffect(() => {
    return () => {
      if (highlightTimerRef.current != null) {
        window.clearTimeout(highlightTimerRef.current);
      }
    };
  }, []);

  return {
    highlightedMessageId,
    matchingMessageIds,
    normalizedSearchQuery,
    scrollToMessage,
    submitSearch,
  };
}
