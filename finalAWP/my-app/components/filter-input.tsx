"use client";

import { useState } from "react";
import { Search, X } from "lucide-react";

type StockDatum = {
  symbol: string;
  label: string;
  price: number;
  open: number;
  high: number;
  low: number;
  previousClose: number;
  changePercent: number;
  volume: number;
  marketCap: number;
};

type FilterInputProps = {
  initialData: StockDatum[];
};

export default function FilterInput({ initialData }: FilterInputProps) {
  const [query, setQuery] = useState("");
  const [focused, setFocused] = useState(false);

  const filtered = query.trim() === ""
    ? initialData
    : initialData.filter((s) =>
        s.symbol.toLowerCase().includes(query.trim().toLowerCase()) ||
        s.label.toLowerCase().includes(query.trim().toLowerCase())
      );

  return (
    <div className="w-full max-w-md">
      <label htmlFor="stock-filter" className="sr-only">
        Filter stocks by ticker or name
      </label>
      <div
        className={`
          relative flex items-center gap-2 px-3.5 py-2.5 rounded-lg border transition-all duration-200
          ${focused
            ? "border-zinc-600 bg-zinc-900/60 ring-1 ring-zinc-700"
            : "border-zinc-800 bg-zinc-900/40"
          }
        `}
      >
        <Search
          className={`
            w-4 h-4 transition-colors shrink-0
            ${focused ? "text-zinc-400" : "text-zinc-600"}
          `}
        />
        <input
          id="stock-filter"
          type="text"
          placeholder="Filter by name or ticker…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onFocus={() => setFocused(true)}
          onBlur={() => setFocused(false)}
          className="
            bg-transparent text-sm text-zinc-200 placeholder-zinc-600
            outline-none w-full font-medium
          "
        />
        {query && (
          <button
            type="button"
            onClick={() => setQuery("")}
            className="text-zinc-600 hover:text-zinc-300 transition-colors shrink-0"
            aria-label="Clear filter"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        )}
      </div>
      {query && filtered.length === 0 && (
        <p className="text-xs text-zinc-600 mt-2 px-1">
          No results for &ldquo;{query}&rdquo;
        </p>
      )}
      {filtered.length < initialData.length && query && (
        <p className="text-xs text-zinc-600 mt-2 px-1">
          Showing {filtered.length} of {initialData.length}
        </p>
      )}
    </div>
  );
}
