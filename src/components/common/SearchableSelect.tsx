import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { Check, ChevronDown, Search } from "lucide-react";

export type Option = { value: string; label: string };

type Props = {
  value: string;
  onChange: (value: string) => void;
  options: Option[];
  placeholder: string; // e.g. "All Districts"
  className?: string;
  searchable?: boolean;
    allValue?: string; // "All ..." option select korle ei value jabe (default "")
  compact?: boolean; // chhoto button // false dile search bar thakbe na (default true)
};

const SEARCH_BOX_H = 34; // search bar-er height
const TOP_OFFSET = 70; // window scroll hole: screen-er upor theke button koto px niche thakbe (fixed navbar-er height)
const MIN_LIST_H = 120;
const MAX_LIST_H = 340;
const PANEL_MIN_W = 190;

type Pos = {
  left: number;
  width: number;
  top?: number;
  bottom?: number;
};

// button-er nearest scrollable parent (modal etc.). Na thakle null = window
function getScrollParent(el: HTMLElement | null): HTMLElement | null {
  let p = el?.parentElement ?? null;
  while (p && p !== document.body) {
    const oy = getComputedStyle(p).overflowY;
    if (oy === "auto" || oy === "scroll") return p;
    p = p.parentElement;
  }
  return null;
}

export default function SearchableSelect({
  value,
  onChange,
  options,
  placeholder,
  className = "",
  searchable = true,
  allValue = "",
  compact = false,
}: Props) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [listMaxH, setListMaxH] = useState(MAX_LIST_H);
  const [pos, setPos] = useState<Pos | null>(null);

  const wrapRef = useRef<HTMLDivElement>(null);
  const btnRef = useRef<HTMLButtonElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const searchingRef = useRef(false);
  const timersRef = useRef<ReturnType<typeof setTimeout>[]>([]);
  const padTargetRef = useRef<HTMLElement | null>(null);
  const prevPaddingRef = useRef("");

  const searchH = searchable ? SEARCH_BOX_H : 0;

  const selectedLabel = options.find((o) => o.value === value)?.label;
  const hasValue = value !== allValue;

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return options;
    return options.filter((o) => o.label.toLowerCase().includes(q));
  }, [options, query]);

  // shudhu position/height hishab kore (scroll kore na)
  const reposition = useCallback(() => {
    const btn = btnRef.current;
    if (!btn) return;
    const vv = window.visualViewport;
    const viewH = vv ? vv.height : window.innerHeight;
    const viewW = vv ? vv.width : window.innerWidth;
    const offsetTop = vv ? vv.offsetTop : 0;
    const rect = btn.getBoundingClientRect();

    const width = Math.max(rect.width, PANEL_MIN_W);
    let left = rect.left;
    if (left + width > viewW - 4) left = Math.max(4, viewW - 4 - width);

    const below = offsetTop + viewH - rect.bottom - searchH - 12;
    const above = rect.top - offsetTop - searchH - 12;

    // niche jaygа kom + upore beshi (ar search mode na) hole upor dike khule
    const flip = !searchingRef.current && below < MIN_LIST_H && above > below;

    if (flip) {
      setPos({ left, width, bottom: window.innerHeight - rect.top + 4 });
      setListMaxH(Math.max(MIN_LIST_H, Math.min(MAX_LIST_H, above)));
    } else {
      setPos({ left, width, top: rect.bottom + 4 });
      setListMaxH(Math.max(MIN_LIST_H, Math.min(MAX_LIST_H, below)));
    }
  }, [searchH]);

  // search mode: button ke upore tule ane, tarpor position hishab kore
  const layout = useCallback(() => {
    const btn = btnRef.current;
    if (!btn) return;

    if (searchingRef.current) {
      const vv = window.visualViewport;
      const offsetTop = vv ? vv.offsetTop : 0;
      const sp = getScrollParent(btn);
      const rect = btn.getBoundingClientRect();

      if (sp) {
        const target = Math.max(sp.getBoundingClientRect().top, offsetTop) + 8;
        const delta = rect.top - target;
        if (Math.abs(delta) > 2) sp.scrollTop += delta;
      } else {
        const delta = rect.top - offsetTop - TOP_OFFSET;
        if (Math.abs(delta) > 2) window.scrollBy({ top: delta, behavior: "auto" });
      }

      const r2 = btn.getBoundingClientRect();
      const viewH = vv ? vv.height : window.innerHeight;
      const below = offsetTop + viewH - r2.bottom - searchH - 12;
      setListMaxH(Math.max(MIN_LIST_H, Math.min(MAX_LIST_H, below)));
    }
    reposition();
  }, [reposition, searchH]);

  const clearTimers = () => {
    timersRef.current.forEach(clearTimeout);
    timersRef.current = [];
  };

  const restorePadding = () => {
    if (padTargetRef.current) {
      padTargetRef.current.style.paddingBottom = prevPaddingRef.current;
      padTargetRef.current = null;
    }
  };

  const handleSearchFocus = () => {
    searchingRef.current = true;
    // scroll korar jonno temporary extra jaygа (modal / page duto-tei)
    if (!padTargetRef.current) {
      const target = getScrollParent(btnRef.current) ?? document.body;
      padTargetRef.current = target;
      prevPaddingRef.current = target.style.paddingBottom;
      target.style.paddingBottom = "70vh";
    }
    layout();
    clearTimers();
    timersRef.current.push(setTimeout(layout, 150));
    timersRef.current.push(setTimeout(layout, 400)); // keyboard animation sesh hole abar
  };

  const closeDropdown = useCallback(() => {
    clearTimers();
    restorePadding();
    searchingRef.current = false;
    setOpen(false);
  }, []);

  // open hole (keyboard khulbe na)
  useEffect(() => {
    if (!open) return;
    setQuery("");
    searchingRef.current = false;
    reposition();

    const vv = window.visualViewport;
    const onScroll = () => reposition();
    vv?.addEventListener("resize", layout);
    window.addEventListener("resize", layout);
    window.addEventListener("scroll", onScroll, true); // modal-er scroll-o dhore
    return () => {
      vv?.removeEventListener("resize", layout);
      window.removeEventListener("resize", layout);
      window.removeEventListener("scroll", onScroll, true);
    };
  }, [open, layout, reposition]);

  useEffect(() => {
    return () => {
      clearTimers();
      restorePadding();
    };
  }, []);

  // bahire click/touch korle close
  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent | TouchEvent) => {
      const t = e.target as Node;
      if (wrapRef.current?.contains(t) || panelRef.current?.contains(t)) return;
      closeDropdown();
    };
    document.addEventListener("mousedown", handler);
    document.addEventListener("touchstart", handler);
    return () => {
      document.removeEventListener("mousedown", handler);
      document.removeEventListener("touchstart", handler);
    };
  }, [open, closeDropdown]);

  const select = (v: string) => {
    onChange(v);
    closeDropdown();
  };

  return (
    <div ref={wrapRef} className={`relative ${className}`}>
      <style>{`
        @keyframes ssPop {
          from { opacity: 0; transform: translateY(-4px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        .ss-pop { animation: ssPop 120ms ease-out; }
        .ss-scroll::-webkit-scrollbar { width: 5px; }
        .ss-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 9999px; }
        .ss-scroll::-webkit-scrollbar-track { background: transparent; }
      `}</style>

      {/* Trigger button */}
      <button
        ref={btnRef}
        type="button"
        onClick={() => (open ? closeDropdown() : setOpen(true))}
        className={`flex w-full items-center justify-between gap-1 rounded-lg border shadow-sm outline-none transition-all duration-150 ${
  compact ? "px-1.5 py-1 text-[11px]" : "px-2 py-1.5 text-xs"
} ${
  open
    ? "border-blue-500 bg-white ring-2 ring-blue-500/20"
    : hasValue
    ? "border-slate-300 bg-white font-medium text-slate-800 hover:border-slate-400"
    : "border-slate-300 bg-white text-slate-600 hover:border-slate-400"
}`}
      >
        <span className="truncate">{selectedLabel || placeholder}</span>
        <ChevronDown
          size={13}
          className={`shrink-0 transition-transform duration-200 ${
            open ? "rotate-180 text-blue-600" : hasValue ? "text-blue-500" : "text-slate-400"
          }`}
        />
      </button>

      {/* Dropdown panel (portal: modal-er bhitore kete jabe na) */}
      {open &&
        pos &&
        createPortal(
          <div
            ref={panelRef}
            className="ss-pop fixed z-[3000] overflow-hidden rounded-xl border border-slate-200 bg-white shadow-[0_10px_30px_-8px_rgba(15,23,42,0.25)] ring-1 ring-black/5"
            style={{
              left: pos.left,
              width: pos.width,
              top: pos.top,
              bottom: pos.bottom,
            }}
          >
            {/* Search bar (searchable=false hole thakbe na) */}
            {searchable && (
              <div
                className="flex items-center border-b border-slate-100 px-1.5"
                style={{ height: SEARCH_BOX_H }}
              >
                <div className="flex w-full items-center gap-1.5 rounded-md bg-slate-100 px-2 py-0.5 transition focus-within:bg-white focus-within:ring-1 focus-within:ring-blue-500">
                  <Search size={12} className="shrink-0 text-slate-400" />
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    onFocus={handleSearchFocus}
                    placeholder="Search..."
                    className="w-full bg-transparent py-0 text-base text-slate-800 outline-none placeholder:text-slate-400 sm:text-xs"
                  />
                </div>
              </div>
            )}

            {/* Options */}
            <ul
              className="ss-scroll overflow-y-auto overscroll-contain p-1"
              style={{ maxHeight: listMaxH }}
            >
              {/* Search korar somoy ei option lukano thakbe */}
              {!query.trim() && (
                <li>
                    <button
                    type="button"
                    onClick={() => select(allValue)}
                    className={`flex w-full items-center justify-between rounded-md px-2 py-1.5 text-left text-xs transition ${
                        value === allValue
                        ? "bg-blue-50 font-semibold text-blue-700"
                        : "text-slate-500 hover:bg-slate-100"
                    }`}
                    >
                    <span>{placeholder}</span>
                    {value === allValue && <Check size={12} />}
                    </button>
                </li>
                )}

              {filtered.map((o) => {
                const active = value === o.value;
                return (
                  <li key={o.value}>
                    <button
                      type="button"
                      onClick={() => select(o.value)}
                      className={`flex w-full items-center justify-between gap-2 rounded-md px-2 py-1.5 text-left text-xs transition ${
                        active
                          ? "bg-blue-50 font-semibold text-blue-700"
                          : "text-slate-700 hover:bg-blue-50 hover:text-blue-700 active:bg-blue-100"
                      }`}
                    >
                      <span className="truncate">{o.label}</span>
                      {active && <Check size={12} className="shrink-0" />}
                    </button>
                  </li>
                );
              })}

              {filtered.length === 0 && (
                <li className="px-2 py-3 text-center text-xs text-slate-400">
                  No results found
                </li>
              )}
            </ul>
          </div>,
          document.body
        )}
    </div>
  );
}