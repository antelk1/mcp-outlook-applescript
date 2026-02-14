/**
 * Pagination envelope for list/search tool responses.
 *
 * Every list/search tool returns `{ items, count, hasMore }` so that
 * LLM callers know whether more results exist and can paginate.
 */
/**
 * Trims an over-fetched result set and produces a PaginatedResult.
 *
 * Callers should pass `limit + 1` to their data source. If `limit + 1`
 * results come back, `hasMore` is true and the extra item is trimmed.
 */
export function paginate(items, limit) {
    const hasMore = items.length > limit;
    const trimmed = hasMore ? items.slice(0, limit) : items;
    return { items: trimmed, count: trimmed.length, hasMore };
}
