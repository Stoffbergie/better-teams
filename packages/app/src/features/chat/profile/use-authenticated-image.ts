import { teamsProfileService } from "@better-teams/app/services/teams/profile";
import { useQuery } from "@tanstack/react-query";

/**
 * React hook that fetches an authenticated image from Teams AMS/CDN.
 * Returns a local asset URL (or null while loading / on failure).
 */
export function useAuthenticatedImage(
  imageUrl: string | undefined,
  tenantId: string | undefined | null,
): { src: string | null; loading: boolean; error: boolean } {
  const queryImageUrl = imageUrl ?? null;
  const queryTenantId = tenantId ?? null;
  const { data, isError, isFetched, isLoading } = useQuery({
    queryKey: ["authenticated-image", queryTenantId, queryImageUrl],
    enabled: Boolean(queryImageUrl && queryTenantId),
    staleTime: 30 * 60_000,
    gcTime: 2 * 60 * 60_000,
    retry: false,
    queryFn: async () => {
      if (!queryImageUrl || !queryTenantId) return null;
      return teamsProfileService.fetchAuthenticatedImageSrc(
        queryTenantId,
        queryImageUrl,
      );
    },
  });

  return {
    src: data ?? null,
    loading: isLoading,
    error: isError || (isFetched && !data),
  };
}
