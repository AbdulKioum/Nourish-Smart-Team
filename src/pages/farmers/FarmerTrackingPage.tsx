import React, { useEffect, useMemo, useState } from "react";
import { supabase } from "../../lib/supabase";
import { getFarmers } from "../../services/farmerService";
import { useAuth } from "../../context/AuthContext";
import {
  MapContainer,
  Marker,
  Polyline,
  CircleMarker,
  useMap,
} from "react-leaflet";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import MapLayers from "../../components/maps/MapLayers";
import {
  Search,
  Send,
  X,
  MapPin,
  Phone,
  RotateCcw,
  CheckCircle2,
  Plus,
  ChevronDown,
  ChevronUp,
  User,
  Crosshair,
} from "lucide-react";
import { useSearchParams } from "react-router-dom";

/* -------------------------------------------------------------------------- */
/*  Helpers                                                                    */
/* -------------------------------------------------------------------------- */

const POULTRY_FIELDS = [
  "broiler",
  "layer",
  "sonali",
  "duck",
  "quail",
  "deshi_chicken",
];
const CATTLE_FIELDS = ["dairy_cow", "bull", "calf_heifer"];
const FISH_FIELDS = [
  "pangus",
  "tilapia",
  "koi",
  "shing",
  "pabda",
  "gulsha",
  "magur",
  "carp",
  "golda",
  "bagda",
  "tengra",
];

function parseList(val: any): string[] {
  if (!val) return [];
  if (Array.isArray(val)) return val.map(String).filter(Boolean);
  if (typeof val === "string") {
    try {
      const parsed = JSON.parse(val);
      if (Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
    } catch {
      return val
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    }
  }
  return [];
}

function getFarmerBadgeInfo(farmer: any) {
  if (farmer.farm_category === "Poultry") {
    if (Number(farmer.layer || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🥚", label: "Layer", border: "#b45309" };
    }
    if (Number(farmer.sonali || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🐓", label: "Sonali", border: "#c2410c" };
    }
    if (Number(farmer.duck || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🦆", label: "Duck", border: "#0f766e" };
    }
    return { emoji: "🐔", label: "Broiler", border: "#d97706" };
  }
  if (farmer.farm_category === "Cattle") {
    return { emoji: "🐄", label: "Cattle", border: "#15803d" };
  }
  if (farmer.farm_category === "Fish") {
    return { emoji: "🐟", label: "Fish", border: "#0369a1" };
  }
  return { emoji: "🏡", label: "Farm", border: "#334155" };
}

function getStockBreakdownList(farmer: any): { label: string; qty: number }[] {
  const list: { label: string; qty: number }[] = [];
  [...POULTRY_FIELDS, ...CATTLE_FIELDS, ...FISH_FIELDS].forEach((k) => {
    const val = Number(farmer?.[k] || 0);
    if (val > 0) {
      const label = k
        .replace(/_/g, " ")
        .replace(/\b\w/g, (c) => c.toUpperCase());
      list.push({ label, qty: val });
    }
  });
  return list;
}

function esc(s: string) {
  return (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}


function buildMarker(farmer: any, highlight = false, distText?: string) {
  const { emoji, border } = getFarmerBadgeInfo(farmer);
  const name = esc((farmer.farmer_name || "Farmer").slice(0, 16));
  const ring = highlight
    ? "box-shadow:0 0 0 4px rgba(16,185,129,0.35);"
    : "box-shadow:0 1px 2px rgba(0,0,0,0.15);";
  const distHtml = distText
    ? `<div style="margin-top:2px;background:#0f172a;color:#fff;font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;white-space:nowrap;">${esc(
        distText
      )}</div>`
    : "";
  const html = `
    <div style="display:flex;flex-direction:column;align-items:center;transform:translateY(-10px);">
      <span style="font-size:${highlight ? 30 : 24}px;line-height:1;filter:drop-shadow(0 2px 3px rgba(0,0,0,0.35));">${emoji}</span>
      <div style="margin-top:2px;background:rgba(255,255,255,0.96);border:1px solid ${border};color:#0f172a;font-size:10px;font-weight:700;padding:0 5px;border-radius:4px;white-space:nowrap;${ring}">
        ${name}
      </div>
      ${distHtml}
    </div>`;
  return L.divIcon({
    className: "ft-div-icon",
    html,
    iconSize: [70, 56],
    iconAnchor: [35, 42],
  });
}

function toYmd(iso: string) {
  const d = new Date(iso);
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${m}-${day}`;
}

function fmtDateTime(iso: string) {
  return new Date(iso).toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });
}

function timeAgo(iso: string) {
  const diff = Date.now() - new Date(iso).getTime();
  const min = Math.floor(diff / 60000);
  if (min < 1) return "just now";
  if (min < 60) return `${min}m ago`;
  const hr = Math.floor(min / 60);
  if (hr < 24) return `${hr}h ago`;
  const day = Math.floor(hr / 24);
  if (day < 30) return `${day}d ago`;
  const mon = Math.floor(day / 30);
  return mon < 12 ? `${mon}mo ago` : `${Math.floor(mon / 12)}y ago`;
}

function uniqSorted(list: string[]) {
  return Array.from(new Set(list.filter(Boolean))).sort((a, b) =>
    a.localeCompare(b)
  );
}

function calcDistanceKm(
  lat1: number,
  lon1: number,
  lat2: number,
  lon2: number
): number {
  const R = 6371;
  const dLat = ((lat2 - lat1) * Math.PI) / 180;
  const dLon = ((lon2 - lon1) * Math.PI) / 180;
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos((lat1 * Math.PI) / 180) *
      Math.cos((lat2 * Math.PI) / 180) *
      Math.sin(dLon / 2) *
      Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function formatDistance(km: number): string {
  if (km < 1) return `${Math.round(km * 1000)} m`;
  return `${km.toFixed(2)} km`;
}

function FlyToLocation({ target }: { target: [number, number] | null }) {
  const map = useMap();
  useEffect(() => {
    if (target) {
      map.flyTo(target, 15, { animate: true, duration: 0.8 });
    }
  }, [target, map]);
  return null;
}

function FitBounds({
  positions,
  flyTarget,
}: {
  positions: [number, number][];
  flyTarget?: { pos: [number, number]; zoom?: number; ts: number } | null;
}) {
  const map = useMap();

  useEffect(() => {
    if (flyTarget) {
      map.flyTo(flyTarget.pos, flyTarget.zoom || 15, { duration: 0.8 });
    }
  }, [flyTarget, map]);

  useEffect(() => {
    if (flyTarget) return;
    if (positions.length === 0) return;
    if (positions.length === 1) {
      map.setView(positions[0], 15);
      return;
    }
    map.fitBounds(positions, { padding: [45, 45] });
  }, [positions, map, flyTarget]);

  return null;
}

const selectCls =
  "w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs text-slate-800 outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500";
const labelCls = "block text-[11px] font-semibold text-slate-500 mb-0.5";

/* -------------------------------------------------------------------------- */
/*  Shared Components                                                          */
/* -------------------------------------------------------------------------- */

function NoteThread({
  notes,
  initialLimit = 10,
}: {
  notes: any[];
  initialLimit?: number;
}) {
  const [limit, setLimit] = useState(initialLimit);

  if (notes.length === 0) {
    return (
      <div className="rounded-lg border border-dashed border-slate-300 bg-slate-50 p-4 text-center text-xs text-slate-400">
        No tracking messages yet.
      </div>
    );
  }

  const visibleNotes = notes.slice(0, limit);
  const remaining = notes.length - visibleNotes.length;

  return (
    <div className="space-y-2">
      {visibleNotes.map((n, i) => (
        <div
          key={n.id}
          id={`track-note-${n.id}`}
          className={`rounded-lg border px-3 py-2 ${
            i === 0
              ? "border-emerald-400 bg-emerald-50/80 shadow-2xs"
              : "border-slate-200 bg-white"
          }`}
        >
          <div className="flex flex-wrap items-center justify-between gap-1.5 text-[11px]">
            <span className="flex items-center gap-1.5 font-bold text-slate-800">
              <span
                className={`flex h-4.5 w-4.5 items-center justify-center rounded-full text-[9px] font-bold text-white ${
                  i === 0 ? "bg-emerald-600" : "bg-slate-700"
                }`}
              >
                {(n.user_name || "?").charAt(0).toUpperCase()}
              </span>
              <span>{n.user_name || "Officer"}</span>
              {i === 0 && (
                <span className="rounded bg-emerald-600 px-1.5 py-px text-[9px] font-bold text-white uppercase tracking-wide">
                  Latest
                </span>
              )}
            </span>
            <span className="text-slate-500 font-medium">
              {fmtDateTime(n.created_at)} ({timeAgo(n.created_at)})
            </span>
          </div>
          <p className="mt-1 whitespace-pre-wrap break-words text-[13px] sm:text-sm font-medium leading-snug text-slate-900">
            {n.note}
          </p>
        </div>
      ))}

      {remaining > 0 && (
        <button
          type="button"
          onClick={() => setLimit((prev) => prev + 10)}
          className="w-full rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 py-1.5 text-center text-xs font-semibold text-emerald-700 transition"
        >
          Load more messages ({remaining} older)
        </button>
      )}
    </div>
  );
}

function FarmerInfoGrid({ farmer }: { farmer: any }) {
  const stock = getStockBreakdownList(farmer);
  const address = [
    farmer.village,
    farmer.union_name,
    farmer.upazila,
    farmer.district,
  ]
    .filter(Boolean)
    .join(", ");

  const rows: { label: string; value: React.ReactNode }[] = [
    { label: "Company", value: farmer.companies?.company_name || "-" },
    { label: "Agent", value: farmer.agents?.agent_name || "-" },
    {
      label: "Mobile",
      value: farmer.mobile ? (
        <a
          href={`tel:${farmer.mobile}`}
          className="inline-flex items-center gap-1 font-semibold text-emerald-700 hover:underline"
        >
          <Phone size={11} /> {farmer.mobile}
        </a>
      ) : (
        "-"
      ),
    },
    { label: "Address", value: address || "-" },
    {
      label: "Present Feed",
      value:
        farmer.last_feed_company ||
        farmer.feed_type ||
        farmer.companies?.company_name ||
        "-",
    },
    { label: "Status", value: farmer.status || "Active" },
  ];

  return (
    <div className="space-y-2 text-xs">
      <div className="grid grid-cols-[88px_1fr] gap-x-2 gap-y-1">
        {rows.map((r) => (
          <React.Fragment key={r.label}>
            <span className="text-slate-500">{r.label}</span>
            <span className="text-slate-900 font-medium">{r.value}</span>
          </React.Fragment>
        ))}
      </div>
      <div>
        <div className="mb-1 text-slate-500">
          {farmer.farm_category || "Farm"} stock
        </div>
        {stock.length === 0 ? (
          <span className="text-slate-400">No stock recorded</span>
        ) : (
          <div className="flex flex-wrap gap-1">
            {stock.map((s) => (
              <span
                key={s.label}
                className="rounded-md border border-slate-200 bg-white px-1.5 py-0.5 font-semibold text-slate-800"
              >
                {s.label}: <span className="text-blue-700">{s.qty}</span>
              </span>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function MapPickerModal({
  farmers,
  initDistrict,
  initUpazila,
  initAgent,
  onPick,
  onClose,
}: {
  farmers: any[];
  initDistrict: string;
  initUpazila: string;
  initAgent: string;
  onPick: (f: any) => void;
  onClose: () => void;
}) {
  const [d, setD] = useState(initDistrict);
  const [u, setU] = useState(initUpazila);
  const [a, setA] = useState(initAgent);
  const [myLoc, setMyLoc] = useState<[number, number] | null>(null);
  const [flyTarget, setFlyTarget] = useState<[number, number] | null>(null);
  const [locating, setLocating] = useState(false);

  const districts = useMemo(
    () => uniqSorted(farmers.map((f) => f.district)),
    [farmers]
  );
  const upazilas = useMemo(
    () =>
      uniqSorted(
        farmers.filter((f) => !d || f.district === d).map((f) => f.upazila)
      ),
    [farmers, d]
  );
  const agentOptions = useMemo(() => {
    const map = new Map<string, string>();
    farmers
      .filter(
        (f) => (!d || f.district === d) && (!u || f.upazila === u) && f.agent_id
      )
      .forEach((f) =>
        map.set(String(f.agent_id), f.agents?.agent_name || "Unnamed agent")
      );
    return Array.from(map.entries()).sort((x, y) => x[1].localeCompare(y[1]));
  }, [farmers, d, u]);

  const shown = useMemo(
    () =>
      farmers.filter(
        (f) =>
          (!d || f.district === d) &&
          (!u || f.upazila === u) &&
          (!a || String(f.agent_id) === a)
      ),
    [farmers, d, u, a]
  );
  const withLoc = useMemo(
    () => shown.filter((f) => f.latitude && f.longitude),
    [shown]
  );
  const positions = useMemo(
    () =>
      withLoc.map(
        (f) => [Number(f.latitude), Number(f.longitude)] as [number, number]
      ),
    [withLoc]
  );

  function handleLocateMe() {
    if (!navigator.geolocation) {
      alert("Geolocation is not supported by your browser.");
      return;
    }
    setLocating(true);
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        const coords: [number, number] = [
          pos.coords.latitude,
          pos.coords.longitude,
        ];
        setMyLoc(coords);
        setFlyTarget([...coords]);
        setLocating(false);
      },
      () => {
        setLocating(false);
        alert("Unable to retrieve your location. Please allow GPS access.");
      },
      { enableHighAccuracy: true, timeout: 15000 }
    );
  }

  return (
    <div className="fixed inset-0 z-[2000] flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="flex h-[88vh] w-full max-w-5xl flex-col overflow-hidden rounded-xl border border-slate-300 bg-white shadow-2xl">
        <div className="flex items-center justify-between border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <div>
            <h2 className="text-sm font-bold text-slate-800">
              Pick a farmer from the map
            </h2>
            <p className="text-[11px] text-slate-500">
              {withLoc.length} on map
              {shown.length - withLoc.length > 0 &&
                `, ${shown.length - withLoc.length} without location`}
              . Tap a farmer to select.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleLocateMe}
              disabled={locating}
              className="flex items-center gap-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 px-3 py-1.5 text-xs font-bold text-white shadow-sm disabled:opacity-60"
            >
              <Crosshair size={13} />
              {locating ? "Locating..." : "Locate Me"}
            </button>
            <button
              onClick={onClose}
              className="text-slate-500 hover:text-slate-800"
            >
              <X size={18} />
            </button>
          </div>
        </div>

        <div className="grid grid-cols-3 gap-2 border-b border-slate-200 px-4 py-2">
          <div>
            <label className={labelCls}>District</label>
            <select
              className={selectCls}
              value={d}
              onChange={(e) => {
                setD(e.target.value);
                setU("");
                setA("");
              }}
            >
              <option value="">All</option>
              {districts.map((x) => (
                <option key={x} value={x}>
                  {x}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className={labelCls}>Upazila</label>
            <select
              className={selectCls}
              value={u}
              onChange={(e) => {
                setU(e.target.value);
                setA("");
              }}
            >
              <option value="">All</option>
              {upazilas.map((x) => (
                <option key={x} value={x}>
                  {x}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className={labelCls}>Agent</label>
            <select
              className={selectCls}
              value={a}
              onChange={(e) => setA(e.target.value)}
            >
              <option value="">All</option>
              {agentOptions.map(([id, name]) => (
                <option key={id} value={id}>
                  {name}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="relative min-h-0 flex-1">
          {/* ম্যাপের ওপর ভাসমান Locate Me বাটন */}
          <button
            type="button"
            onClick={handleLocateMe}
            disabled={locating}
            className="absolute top-3 right-3 z-[1000] flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white/95 hover:bg-blue-50 px-3 py-2 text-xs font-bold text-blue-700 shadow-lg backdrop-blur-xs"
          >
            <Crosshair size={14} className="text-blue-600" />
            {locating ? "Locating..." : "Locate Me"}
          </button>

          <MapContainer
            center={[24.7471, 90.4203]}
            zoom={8}
            style={{ height: "100%", width: "100%", zIndex: 1 }}
          >
            <MapLayers />
            <FitBounds positions={positions} />
            <FlyToLocation target={flyTarget} />
            {myLoc && (
              <CircleMarker
                center={myLoc}
                radius={9}
                pathOptions={{
                  color: "#ffffff",
                  weight: 3,
                  fillColor: "#2563eb",
                  fillOpacity: 1,
                }}
              />
            )}
            {withLoc.map((f) => {
              const distText = myLoc
                ? formatDistance(
                    calcDistanceKm(
                      myLoc[0],
                      myLoc[1],
                      Number(f.latitude),
                      Number(f.longitude)
                    )
                  )
                : undefined;
              return (
                <Marker
                  key={f.id}
                  position={[Number(f.latitude), Number(f.longitude)]}
                  icon={buildMarker(f, false, distText)}
                  eventHandlers={{ click: () => onPick(f) }}
                />
              );
            })}
          </MapContainer>
        </div>
      </div>
    </div>
  );
}

function FarmerMapModal({
  farmer,
  onClose,
  onDetails,
}: {
  farmer: any;
  onClose: () => void;
  onDetails: () => void;
}) {
  const pos: [number, number] = [
    Number(farmer.latitude),
    Number(farmer.longitude),
  ];
  const [myLoc, setMyLoc] = useState<[number, number] | null>(null);
  const [locating, setLocating] = useState(false);

  const distanceKm = useMemo(() => {
    if (!myLoc) return null;
    return calcDistanceKm(myLoc[0], myLoc[1], pos[0], pos[1]);
  }, [myLoc, pos]);

  const boundsPositions = useMemo<[number, number][]>(
    () => (myLoc ? [myLoc, pos] : [pos]),
    [myLoc, pos]
  );

  function handleLocateMe() {
    if (!navigator.geolocation) {
      alert("Geolocation is not supported by your browser.");
      return;
    }
    setLocating(true);
    navigator.geolocation.getCurrentPosition(
      (p) => {
        setMyLoc([p.coords.latitude, p.coords.longitude]);
        setLocating(false);
      },
      () => {
        setLocating(false);
        alert("Unable to retrieve your location. Please allow GPS access.");
      },
      { enableHighAccuracy: true, timeout: 15000 }
    );
  }

  return (
    <div className="fixed inset-0 z-[2000] flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="flex h-[80vh] w-full max-w-3xl flex-col overflow-hidden rounded-xl border border-slate-300 bg-white shadow-2xl">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <div>
            <h2 className="text-sm font-bold text-slate-800">
              {farmer.farmer_name}
            </h2>
            <p className="text-[11px] text-slate-500">
              {[
                farmer.village,
                farmer.union_name,
                farmer.upazila,
                farmer.district,
              ]
                .filter(Boolean)
                .join(", ")}
            </p>
            {distanceKm !== null && (
              <span className="mt-1 inline-block rounded-md bg-emerald-100 border border-emerald-300 px-2 py-0.5 text-[11px] font-bold text-emerald-900">
                Distance from you: {formatDistance(distanceKm)}
              </span>
            )}
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleLocateMe}
              disabled={locating}
              className="flex items-center gap-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 px-3 py-1.5 text-xs font-bold text-white shadow-sm disabled:opacity-60"
            >
              <Crosshair size={13} />
              {locating ? "Locating..." : "Locate Me"}
            </button>
            <button
              onClick={onDetails}
              className="rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
            >
              View details
            </button>
            <button
              onClick={onClose}
              className="text-slate-500 hover:text-slate-800"
            >
              <X size={18} />
            </button>
          </div>
        </div>

        <div className="relative min-h-0 flex-1">
          {/* ম্যাপের ওপর ভাসমান Locate Me ও দূরত্ব ব্যাজ */}
          <div className="absolute top-3 right-3 z-[1000] flex flex-col items-end gap-1.5">
            <button
              type="button"
              onClick={handleLocateMe}
              disabled={locating}
              className="flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white/95 hover:bg-blue-50 px-3 py-2 text-xs font-bold text-blue-700 shadow-lg backdrop-blur-xs"
            >
              <Crosshair size={14} className="text-blue-600" />
              {locating ? "Locating..." : "Locate Me"}
            </button>
            {distanceKm !== null && (
              <div className="rounded-lg bg-slate-900/90 px-2.5 py-1 text-xs font-bold text-white shadow-md">
                📍 {formatDistance(distanceKm)} away
              </div>
            )}
          </div>

          <MapContainer
            center={pos}
            zoom={16}
            style={{ height: "100%", width: "100%", zIndex: 1 }}
          >
            <MapLayers />
            <FitBounds positions={boundsPositions} />
            {myLoc && (
              <>
                <CircleMarker
                  center={myLoc}
                  radius={9}
                  pathOptions={{
                    color: "#ffffff",
                    weight: 3,
                    fillColor: "#2563eb",
                    fillOpacity: 1,
                  }}
                />
                <Polyline
                  positions={[myLoc, pos]}
                  pathOptions={{
                    color: "#2563eb",
                    weight: 3,
                    dashArray: "6 6",
                  }}
                />
              </>
            )}
            <Marker
              position={pos}
              icon={buildMarker(
                farmer,
                true,
                distanceKm !== null
                  ? `${formatDistance(distanceKm)} away`
                  : undefined
              )}
            />
          </MapContainer>
        </div>
      </div>
    </div>
  );
}

function FarmerModal({
  farmer,
  notes,
  initialTab,
  userName,
  onSave,
  onViewMap,
  onClose,
}: {
  farmer: any;
  notes: any[];
  initialTab: "track" | "details";
  userName: string;
  onSave: (farmer: any, text: string) => Promise<boolean>;
  onViewMap: (farmer: any) => void;
  onClose: () => void;
}) {
  const [tab, setTab] = useState<"track" | "details">(initialTab);
  const [text, setText] = useState("");
  const [saving, setSaving] = useState(false);
  const badge = getFarmerBadgeInfo(farmer);

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    if (!text.trim() || saving) return;
    setSaving(true);
    const ok = await onSave(farmer, text);
    setSaving(false);
    if (ok) setText("");
  }

  const tabBtn = (id: "track" | "details", label: string) => (
    <button
      onClick={() => setTab(id)}
      className={`border-b-2 px-3 py-2 text-xs font-semibold ${
        tab === id
          ? "border-emerald-600 text-emerald-700"
          : "border-transparent text-slate-500 hover:text-slate-800"
      }`}
    >
      {label}
    </button>
  );

  const detailRows: [string, React.ReactNode][] = [
    ["Company", farmer.companies?.company_name || "-"],
    ["Agent", farmer.agents?.agent_name || "-"],
    ["Farmer name", farmer.farmer_name],
    ["Mobile", farmer.mobile || "-"],
    [
      "Address",
      [farmer.village, farmer.union_name, farmer.upazila, farmer.district]
        .filter(Boolean)
        .join(", ") || "-",
    ],
    ["Farm category", farmer.farm_category || "-"],
    [
      "Stock / capacity",
      getStockBreakdownList(farmer).length === 0
        ? "-"
        : getStockBreakdownList(farmer)
            .map((b) => `${b.label}: ${b.qty}`)
            .join(", "),
    ],
    [
      "Present Feed",
      farmer.last_feed_company ||
        farmer.feed_type ||
        farmer.companies?.company_name ||
        "-",
    ],
    ["Status", farmer.status || "Active"],
    [
      "Coordinates",
      farmer.latitude && farmer.longitude
        ? `${farmer.latitude}, ${farmer.longitude}`
        : "-",
    ],
    ["Remarks", farmer.remarks || "-"],
  ];

  return (
    <div className="fixed inset-0 z-[2000] flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="flex max-h-[90vh] w-full max-w-xl flex-col overflow-hidden rounded-xl border border-slate-300 bg-white shadow-2xl">
        <div className="flex items-center justify-between gap-3 border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <div className="flex min-w-0 items-center gap-2.5">
            <span className="text-2xl leading-none">{badge.emoji}</span>
            <div className="min-w-0">
              <h2 className="truncate text-sm font-bold text-slate-800">
                {farmer.farmer_name}
              </h2>
              <p className="truncate text-[11px] text-slate-500">
                {farmer.mobile || "No mobile"}
                {farmer.upazila ? `, ${farmer.upazila}` : ""}
              </p>
            </div>
          </div>
          <div className="flex shrink-0 items-center gap-2">
            {farmer.latitude && farmer.longitude && (
              <button
                onClick={() => onViewMap(farmer)}
                className="flex items-center gap-1 rounded-md border border-slate-300 bg-white px-2 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-50"
              >
                <MapPin size={12} /> Map
              </button>
            )}
            <button
              onClick={onClose}
              className="text-slate-500 hover:text-slate-800"
            >
              <X size={18} />
            </button>
          </div>
        </div>

        <div className="flex border-b border-slate-200 px-2">
          {tabBtn("track", `Tracking (${notes.length})`)}
          {tabBtn("details", "Details")}
        </div>

        <div className="overflow-y-auto p-4">
          {tab === "track" ? (
            <div className="space-y-3">
              <form onSubmit={submit} className="space-y-1.5">
                <textarea
                  autoFocus
                  rows={3}
                  value={text}
                  onChange={(e) => setText(e.target.value)}
                  placeholder="Write a new tracking message..."
                  className="w-full resize-none rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500"
                />
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-slate-500">
                    Posting as <b className="text-slate-700">{userName}</b>
                  </span>
                  <button
                    type="submit"
                    disabled={saving || !text.trim()}
                    className="flex items-center gap-1.5 rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700 disabled:opacity-50"
                  >
                    <Send size={12} /> {saving ? "Saving..." : "Save tracking"}
                  </button>
                </div>
              </form>
              <NoteThread notes={notes} initialLimit={10} />
            </div>
          ) : (
            <table className="w-full border-collapse border border-slate-300 text-xs">
              <tbody>
                {detailRows.map(([k, v]) => (
                  <tr key={k}>
                    <td className="w-1/3 border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      {k}
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">{v}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

/* -------------------------------------------------------------------------- */
/*  Farmer Group Card (Shows >= 10 messages per farmer, latest first)          */
/* -------------------------------------------------------------------------- */

function FarmerTrackingGroupCard({
  farmer,
  notes,
  onRetrack,
  onDetails,
  onViewMap,
}: {
  farmer: any;
  notes: any[];
  onRetrack: () => void;
  onDetails: () => void;
  onViewMap: () => void;
}) {
  const badge = getFarmerBadgeInfo(farmer);
  const stockList = getStockBreakdownList(farmer);

  return (
    <div 
    id={`farmer-track-card-${farmer.id}`}
    className="rounded-xl border border-slate-200 bg-white p-3 sm:p-4 shadow-2xs transition hover:border-slate-300">
      <div className="grid grid-cols-1 gap-3 lg:grid-cols-12">
        {/* Farmer Info + Actions (Compact top bar on mobile, left col on desktop) */}
        <div className="flex flex-col justify-between gap-2 border-b border-slate-100 pb-2.5 lg:col-span-4 lg:border-b-0 lg:border-r lg:pb-0 lg:pr-3">
          <div>
            <div className="flex items-start justify-between gap-2">
              <button
                type="button"
                onClick={onDetails}
                className="flex items-center gap-2 text-left min-w-0"
              >
                <span className="text-xl leading-none shrink-0">
                  {badge.emoji}
                </span>
                <span className="truncate text-sm font-bold text-slate-900 hover:text-emerald-700 hover:underline">
                  {farmer.farmer_name}
                </span>
              </button>

              <span className="shrink-0 rounded-full bg-emerald-50 border border-emerald-200 px-2 py-0.5 text-[10px] font-bold text-emerald-800">
                {notes.length} msg{notes.length > 1 ? "s" : ""}
              </span>
            </div>

            <div className="mt-1 flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-slate-600">
              {farmer.mobile && (
                <a
                  href={`tel:${farmer.mobile}`}
                  className="inline-flex items-center gap-1 font-semibold text-emerald-700 hover:underline"
                >
                  <Phone size={10} /> {farmer.mobile}
                </a>
              )}
              {farmer.village && <span>• {farmer.village}</span>}
              {farmer.upazila && (
                <span className="rounded bg-slate-100 px-1.5 py-px text-[10px] font-semibold text-slate-700">
                  {farmer.upazila}
                </span>
              )}
              {farmer.companies?.company_name && (
                <span className="rounded bg-amber-50 border border-amber-200 px-1.5 py-px text-[10px] font-semibold text-amber-800">
                  {farmer.companies.company_name}
                </span>
              )}
            </div>

            {stockList.length > 0 && (
              <div className="mt-1.5 flex flex-wrap gap-1">
                {stockList.map((s) => (
                  <span
                    key={s.label}
                    className="rounded border border-slate-200 bg-slate-50 px-1.5 py-0.5 text-[10px] font-semibold text-slate-700"
                  >
                    {s.label}: <b className="text-blue-700">{s.qty}</b>
                  </span>
                ))}
              </div>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-1.5 pt-1">
            <button
              type="button"
              onClick={onRetrack}
              className="inline-flex flex-1 sm:flex-initial items-center justify-center gap-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 px-3 py-1.5 text-xs font-bold text-white shadow-2xs transition"
            >
              <RotateCcw size={12} /> Retrack
            </button>
            <button
              type="button"
              onClick={onViewMap}
              disabled={!farmer.latitude || !farmer.longitude}
              className="inline-flex flex-1 sm:flex-initial items-center justify-center gap-1 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 px-3 py-1.5 text-xs font-semibold text-slate-700 disabled:opacity-40 transition"
            >
              <MapPin size={12} /> View on Map
            </button>
          </div>
        </div>

        {/* Tracking Messages Thread (Minimum 10 visible, latest on top, no horizontal scroll on mobile) */}
        <div className="lg:col-span-8 min-w-0" id={`farmer-notes-${farmer.id}`}>
          <NoteThread notes={notes} initialLimit={10} />
        </div>
      </div>
    </div>
  );
}

/* -------------------------------------------------------------------------- */
/*  Main Page                                                                  */
/* -------------------------------------------------------------------------- */

export default function FarmerTrackingPage() {
  const { role } = useAuth();
  if (role === "viewer") {
    return <div className="p-6 text-center text-slate-500">Access Denied</div>;
  }
  return <TrackingContent role={role} />;
}

function TrackingContent({ role }: { role: any }) {
  const isFullAccess = role === "super_admin" || role === "manager";

  const [currentUserId, setCurrentUserId] = useState("");
  const [currentUserName, setCurrentUserName] = useState("User");
  const [farmers, setFarmers] = useState<any[]>([]);
  const [trackings, setTrackings] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState("");

  // Collapsible "Track a Farmer" panel state
  const [trackOpen, setTrackOpen] = useState(false);
  const [selDistrict, setSelDistrict] = useState("");
  const [selUpazila, setSelUpazila] = useState("");
  const [selCompany, setSelCompany] = useState("");
  const [selAgent, setSelAgent] = useState("");
  const [farmerQuery, setFarmerQuery] = useState("");
  const [comboOpen, setComboOpen] = useState(false);
  const [pickedFarmer, setPickedFarmer] = useState<any>(null);
  const [noteText, setNoteText] = useState("");
  const [saving, setSaving] = useState(false);
  const [showPicker, setShowPicker] = useState(false);

  // Tracking Log Filters (Default: All Messages)
  const [dateMode, setDateMode] = useState<"all" | "single" | "range">("all");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [logSearch, setLogSearch] = useState("");
  const [logUser, setLogUser] = useState("");
  const [logDistrict, setLogDistrict] = useState("");
  const [logUpazila, setLogUpazila] = useState("");
  const [visibleFarmersCount, setVisibleFarmersCount] = useState(20);

  // Modals
  const [modal, setModal] = useState<{
    farmer: any;
    tab: "track" | "details";
  } | null>(null);
  const [mapFarmer, setMapFarmer] = useState<any>(null);

  
  /* ----------------------------- data loading ----------------------------- */

  async function loadTrackings() {
    const pageSize = 1000;
    let from = 0;
    let all: any[] = [];
    while (true) {
      const { data, error } = await supabase
        .from("farmer_trackings")
        .select("*")
        .order("created_at", { ascending: false })
        .range(from, from + pageSize - 1);
      if (error || !data) break;
      all = all.concat(data);
      if (data.length < pageSize) break;
      from += pageSize;
    }
    setTrackings(all);
  }

  async function loadData() {
    setLoading(true);
    const {
      data: { user },
    } = await supabase.auth.getUser();

    let uDistricts: string[] = [];
    let uUpazilas: string[] = [];
    let uName = user?.email?.split("@")[0] || "Officer";

    if (user) {
      setCurrentUserId(user.id);
      const { data: profile } = await supabase
        .from("users")
        .select("full_name, email, role, districts, upazilas")
        .eq("id", user.id)
        .single();
      if (profile) {
        uName = profile.full_name || profile.email || uName;
        uDistricts = parseList(profile.districts);
        uUpazilas = parseList(profile.upazilas);
      }
    }
    setCurrentUserName(uName);

    const farmerData = (await getFarmers()) || [];
    const allowed = isFullAccess
      ? farmerData
      : farmerData.filter((f: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(f.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(f.district);
          return matchUpa && matchDist;
        });
    setFarmers(allowed);

    const ds = uniqSorted(allowed.map((f: any) => f.district));
    if (ds.length === 1) {
      setSelDistrict(ds[0]);
      const us = uniqSorted(allowed.map((f: any) => f.upazila));
      if (us.length === 1) setSelUpazila(us[0]);
    }

    await loadTrackings();
    setLoading(false);
  }

  useEffect(() => {
    loadData();
  }, [role]);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(""), 2500);
    return () => clearTimeout(t);
  }, [toast]);

  /* ------------------------------ derived data ---------------------------- */

  const farmerMap = useMemo(() => {
    const m = new Map<string, any>();
    farmers.forEach((f) => m.set(String(f.id), f));
    return m;
  }, [farmers]);

  const notesByFarmer = useMemo(() => {
    const m = new Map<string, any[]>();
    trackings.forEach((t) => {
      const k = String(t.farmer_id);
      if (!m.has(k)) m.set(k, []);
      m.get(k)!.push(t);
    });
    return m; // already sorted newest first
  }, [trackings]);

  const selDistricts = useMemo(
    () => uniqSorted(farmers.map((f) => f.district)),
    [farmers]
  );
  const selUpazilas = useMemo(
    () =>
      uniqSorted(
        farmers
          .filter((f) => !selDistrict || f.district === selDistrict)
          .map((f) => f.upazila)
      ),
    [farmers, selDistrict]
  );
  const baseByArea = useMemo(
    () =>
      farmers.filter(
        (f) =>
          (!selDistrict || f.district === selDistrict) &&
          (!selUpazila || f.upazila === selUpazila)
      ),
    [farmers, selDistrict, selUpazila]
  );
  const companyOptions = useMemo(() => {
    const m = new Map<string, string>();
    baseByArea.forEach((f) => {
      if (f.company_id)
        m.set(String(f.company_id), f.companies?.company_name || "Company");
    });
    return Array.from(m.entries()).sort((a, b) => a[1].localeCompare(b[1]));
  }, [baseByArea]);
  const baseByCompany = useMemo(
    () =>
      baseByArea.filter(
        (f) => !selCompany || String(f.company_id) === selCompany
      ),
    [baseByArea, selCompany]
  );
  const agentOptions = useMemo(() => {
    const m = new Map<string, string>();
    baseByCompany.forEach((f) => {
      if (f.agent_id)
        m.set(String(f.agent_id), f.agents?.agent_name || "Unnamed agent");
    });
    return Array.from(m.entries()).sort((a, b) => a[1].localeCompare(b[1]));
  }, [baseByCompany]);
  const farmerChoices = useMemo(() => {
    const q = farmerQuery.trim().toLowerCase();
    return baseByCompany
      .filter((f) => !selAgent || String(f.agent_id) === selAgent)
      .filter(
        (f) =>
          !q ||
          (f.farmer_name || "").toLowerCase().includes(q) ||
          (f.mobile || "").includes(q) ||
          (f.village || "").toLowerCase().includes(q)
      )
      .sort((a, b) =>
        (a.farmer_name || "").localeCompare(b.farmer_name || "")
      );
  }, [baseByCompany, selAgent, farmerQuery]);

  // Unique users who have tracked farmers in the allowed area
  const trackedUsers = useMemo(() => {
    const s = new Set<string>();
    trackings.forEach((t) => {
      if (farmerMap.has(String(t.farmer_id)) && t.user_name) {
        s.add(t.user_name);
      }
    });
    return Array.from(s).sort((a, b) => a.localeCompare(b));
  }, [trackings, farmerMap]);

  const logUpazilas = useMemo(
    () =>
      uniqSorted(
        farmers
          .filter((f) => !logDistrict || f.district === logDistrict)
          .map((f) => f.upazila)
      ),
    [farmers, logDistrict]
  );

  // Group filtered messages per farmer; order farmers by their latest message
  const logGroups = useMemo(() => {
    const q = logSearch.trim().toLowerCase();
    const groups = new Map<string, { farmer: any; notes: any[] }>();

    for (const t of trackings) {
      const key = String(t.farmer_id);
      const f = farmerMap.get(key);
      if (!f) continue;

      if (dateMode === "single" && dateFrom) {
        if (toYmd(t.created_at) !== dateFrom) continue;
      } else if (dateMode === "range") {
        const ymd = toYmd(t.created_at);
        if (dateFrom && ymd < dateFrom) continue;
        if (dateTo && ymd > dateTo) continue;
      }

      if (logUser && t.user_name !== logUser) continue;
      if (logDistrict && f.district !== logDistrict) continue;
      if (logUpazila && f.upazila !== logUpazila) continue;

      if (
        q &&
        !(
          (f.farmer_name || "").toLowerCase().includes(q) ||
          (f.mobile || "").includes(q) ||
          (f.village || "").toLowerCase().includes(q) ||
          (t.note || "").toLowerCase().includes(q) ||
          (t.user_name || "").toLowerCase().includes(q)
        )
      ) {
        continue;
      }

      if (!groups.has(key)) {
        groups.set(key, { farmer: f, notes: [t] });
      } else {
        groups.get(key)!.notes.push(t);
      }
    }

    return Array.from(groups.values());
  }, [
    trackings,
    farmerMap,
    dateMode,
    dateFrom,
    dateTo,
    logSearch,
    logUser,
    logDistrict,
    logUpazila,
  ]);

  useEffect(() => {
    setVisibleFarmersCount(20);
  }, [dateMode, dateFrom, dateTo, logSearch, logUser, logDistrict, logUpazila]);

  /* -------------------------------- actions ------------------------------- */

  async function saveNote(farmer: any, text: string): Promise<boolean> {
    const note = text.trim();
    if (!note) return false;
    const { data, error } = await supabase
      .from("farmer_trackings")
      .insert([
        {
          farmer_id: farmer.id,
          note,
          user_id: currentUserId || null,
          user_name: currentUserName,
        },
      ])
      .select()
      .single();
    if (error) {
      alert(error.message || "Failed to save tracking.");
      return false;
    }
    if (data) setTrackings((prev) => [data, ...prev]);
    else await loadTrackings();
    setToast(`Tracking saved for ${farmer.farmer_name}`);
    return true;
  }

  async function handleSavePicked() {
    if (!pickedFarmer || !noteText.trim() || saving) return;
    setSaving(true);
    const ok = await saveNote(pickedFarmer, noteText);
    setSaving(false);
    if (ok) {
      setNoteText("");
      setPickedFarmer(null);
      setFarmerQuery("");
      setTrackOpen(false);
    }
  }

  function pickFarmer(f: any) {
    setPickedFarmer(f);
    setFarmerQuery(f.farmer_name || "");
    setComboOpen(false);
    setNoteText("");
    setShowPicker(false);
    setTrackOpen(true);
  }

  function clearPicked() {
    setPickedFarmer(null);
    setFarmerQuery("");
    setNoteText("");
  }

  function resetLogFilters() {
    setDateMode("all");
    setDateFrom("");
    setDateTo("");
    setLogSearch("");
    setLogUser("");
    setLogDistrict("");
    setLogUpazila("");
  }

  function openMap(f: any) {
    if (!f.latitude || !f.longitude) {
      alert("No map coordinates saved for this farmer.");
      return;
    }
    setMapFarmer(f);
  }

  const pickedNotes = pickedFarmer
    ? notesByFarmer.get(String(pickedFarmer.id)) || []
    : [];
  const lastPicked = pickedNotes[0];
  const modalNotes = modal
    ? notesByFarmer.get(String(modal.farmer.id)) || []
    : [];
  const pickedBadge = pickedFarmer ? getFarmerBadgeInfo(pickedFarmer) : null;
  const hasLogFilter =
    dateMode !== "all" ||
    dateFrom ||
    dateTo ||
    logSearch ||
    logUser ||
    logDistrict ||
    logUpazila;

    const [searchParams, setSearchParams] = useSearchParams();

  // Realtime subscription on FarmerTrackingPage so new notes appear live here too
  useEffect(() => {
    const ch = supabase
      .channel("page_live_farmer_trackings")
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "farmer_trackings" },
        () => {
          loadTrackings();
        }
      )
      .subscribe();
    return () => {
      supabase.removeChannel(ch);
    };
  }, []);

  // Auto-scroll and highlight when navigated from Navbar Notification
  useEffect(() => {
    if (loading) return;
    const focusFarmerId = searchParams.get("focus");
    const focusNoteId = searchParams.get("note");
    if (!focusFarmerId && !focusNoteId) return;

    const idx = logGroups.findIndex(
      (g) => String(g.farmer.id) === String(focusFarmerId)
    );
    if (idx >= visibleFarmersCount) {
      setVisibleFarmersCount(idx + 5);
    }

    const timer = setTimeout(() => {
      const noteEl = focusNoteId
        ? document.getElementById(`track-note-${focusNoteId}`)
        : null;
      const cardEl = focusFarmerId
        ? document.getElementById(`farmer-track-card-${focusFarmerId}`)
        : null;
      const target = noteEl || cardEl;

      if (target) {
        target.scrollIntoView({ behavior: "smooth", block: "center" });
        target.classList.add("ring-2", "ring-emerald-500", "bg-amber-50");
        setTimeout(() => {
          target.classList.remove("ring-2", "ring-emerald-500", "bg-amber-50");
        }, 3000);
      }

      searchParams.delete("focus");
      searchParams.delete("note");
      setSearchParams(searchParams, { replace: true });
    }, 200);

    return () => clearTimeout(timer);
  }, [loading, logGroups, searchParams]);
  /* --------------------------------- render ------------------------------- */

  return (
    <div className="mx-auto max-w-6xl space-y-3 p-2.5 sm:p-4">
      {/* ---------------- 1. COLLAPSIBLE TRACK A FARMER BUTTON -------------- */}
      <section className="rounded-xl border border-slate-200 bg-white shadow-2xs overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-2 p-2.5 sm:px-4">
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setTrackOpen((o) => !o)}
              className={`inline-flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-bold shadow-xs transition ${
                trackOpen
                  ? "bg-slate-800 text-white"
                  : "bg-emerald-600 hover:bg-emerald-700 text-white"
              }`}
            >
              <Plus size={14} />
              <span>Track a Farmer</span>
              {trackOpen ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </button>

            <button
              type="button"
              onClick={() => setShowPicker(true)}
              className="inline-flex items-center gap-1.5 rounded-lg border border-blue-200 bg-blue-50 hover:bg-blue-100 px-3 py-2 text-xs font-bold text-blue-700 transition"
            >
              <MapPin size={13} /> Pick from Map
            </button>
          </div>

          {pickedFarmer && !trackOpen && (
            <button
              type="button"
              onClick={() => setTrackOpen(true)}
              className="text-xs font-semibold text-emerald-700 hover:underline"
            >
              Selected: {pickedFarmer.farmer_name} (Click to open)
            </button>
          )}
        </div>

        {trackOpen && (
          <div className="border-t border-slate-100 bg-slate-50/40 p-3 sm:p-4 space-y-3">
            <div className="grid grid-cols-2 gap-2 sm:grid-cols-3 md:grid-cols-5">
              <div>
                <label className={labelCls}>District</label>
                <select
                  className={selectCls}
                  value={selDistrict}
                  onChange={(e) => {
                    setSelDistrict(e.target.value);
                    setSelUpazila("");
                    setSelCompany("");
                    setSelAgent("");
                  }}
                >
                  <option value="">All</option>
                  {selDistricts.map((x) => (
                    <option key={x} value={x}>
                      {x}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className={labelCls}>Upazila</label>
                <select
                  className={selectCls}
                  value={selUpazila}
                  onChange={(e) => {
                    setSelUpazila(e.target.value);
                    setSelCompany("");
                    setSelAgent("");
                  }}
                >
                  <option value="">All</option>
                  {selUpazilas.map((x) => (
                    <option key={x} value={x}>
                      {x}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className={labelCls}>Company</label>
                <select
                  className={selectCls}
                  value={selCompany}
                  onChange={(e) => {
                    setSelCompany(e.target.value);
                    setSelAgent("");
                  }}
                >
                  <option value="">All</option>
                  {companyOptions.map(([id, name]) => (
                    <option key={id} value={id}>
                      {name}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className={labelCls}>Agent</label>
                <select
                  className={selectCls}
                  value={selAgent}
                  onChange={(e) => setSelAgent(e.target.value)}
                >
                  <option value="">All</option>
                  {agentOptions.map(([id, name]) => (
                    <option key={id} value={id}>
                      {name}
                    </option>
                  ))}
                </select>
              </div>
              <div className="relative col-span-2 sm:col-span-3 md:col-span-1">
                <label className={labelCls}>
                  Farmer ({farmerChoices.length})
                </label>
                <div className="relative">
                  <Search
                    size={13}
                    className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400"
                  />
                  <input
                    value={farmerQuery}
                    onChange={(e) => {
                      setFarmerQuery(e.target.value);
                      setComboOpen(true);
                    }}
                    onFocus={() => setComboOpen(true)}
                    onBlur={() => setTimeout(() => setComboOpen(false), 140)}
                    placeholder="Search name, mobile, village..."
                    className={`${selectCls} pl-7`}
                  />
                </div>
                {comboOpen && (
                  <div className="absolute right-0 left-0 z-30 mt-1 max-h-64 overflow-y-auto rounded-lg border border-slate-200 bg-white shadow-lg">
                    {farmerChoices.length === 0 ? (
                      <div className="p-3 text-center text-xs text-slate-400">
                        No farmers match these filters.
                      </div>
                    ) : (
                      farmerChoices.slice(0, 40).map((f) => (
                        <button
                          key={f.id}
                          type="button"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            pickFarmer(f);
                          }}
                          className="flex w-full items-center gap-2 border-b border-slate-100 px-3 py-1.5 text-left hover:bg-emerald-50"
                        >
                          <span className="text-lg leading-none">
                            {getFarmerBadgeInfo(f).emoji}
                          </span>
                          <span className="min-w-0">
                            <span className="block truncate text-xs font-semibold text-slate-900">
                              {f.farmer_name}
                            </span>
                            <span className="block truncate text-[11px] text-slate-500">
                              {[f.mobile, f.village, f.companies?.company_name]
                                .filter(Boolean)
                                .join(", ")}
                            </span>
                          </span>
                        </button>
                      ))
                    )}
                  </div>
                )}
              </div>
            </div>

            {pickedFarmer && (
              <div className="grid gap-3 md:grid-cols-5 pt-1">
                <div className="rounded-lg border border-slate-200 bg-white p-3 md:col-span-3">
                  <div className="mb-2 flex items-start justify-between gap-2">
                    <div className="flex min-w-0 items-center gap-2.5">
                      <span className="text-2xl leading-none">
                        {pickedBadge?.emoji}
                      </span>
                      <div className="min-w-0">
                        <div className="truncate text-sm font-bold text-slate-900">
                          {pickedFarmer.farmer_name}
                        </div>
                        <div className="text-[11px] text-slate-500">
                          {pickedBadge?.label}
                          {pickedNotes.length > 0 &&
                            `, tracked ${pickedNotes.length} time${
                              pickedNotes.length > 1 ? "s" : ""
                            }`}
                        </div>
                      </div>
                    </div>
                    <button
                      onClick={clearPicked}
                      title="Clear"
                      className="rounded p-1 text-slate-400 hover:bg-slate-200 hover:text-slate-700"
                    >
                      <X size={14} />
                    </button>
                  </div>
                  <FarmerInfoGrid farmer={pickedFarmer} />
                  {lastPicked && (
                    <div className="mt-2 rounded-md border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-[11px] text-emerald-900">
                      <span className="font-semibold">Last tracking:</span>{" "}
                      {lastPicked.note.length > 90
                        ? lastPicked.note.slice(0, 90) + "..."
                        : lastPicked.note}{" "}
                      <span className="text-emerald-700">
                        ({lastPicked.user_name},{" "}
                        {timeAgo(lastPicked.created_at)})
                      </span>
                    </div>
                  )}
                </div>

                <div className="flex flex-col md:col-span-2">
                  <label className={labelCls}>Tracking message *</label>
                  <textarea
                    value={noteText}
                    onChange={(e) => setNoteText(e.target.value)}
                    placeholder="Write tracking message..."
                    className="min-h-[115px] flex-1 resize-none rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500"
                  />
                  <div className="mt-2 flex items-center justify-between gap-2">
                    <span className="text-[11px] text-slate-500">
                      By <b className="text-slate-700">{currentUserName}</b>
                    </span>
                    <button
                      onClick={handleSavePicked}
                      disabled={saving || !noteText.trim()}
                      className="flex items-center gap-1.5 rounded-lg bg-emerald-600 px-4 py-1.5 text-xs font-bold text-white hover:bg-emerald-700 disabled:opacity-50"
                    >
                      <Send size={12} />{" "}
                      {saving ? "Saving..." : "Save Tracking"}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </section>

      {/* ---------------- 2. MOBILE-FRIENDLY FILTERS & MESSAGES -------------- */}
      <section className="space-y-3">
        {/* Compact Filter Bar */}
        <div className="rounded-xl border border-slate-200 bg-white p-2.5 sm:p-3 shadow-2xs space-y-2">
          <div className="flex flex-wrap items-center justify-between gap-2">
            {/* All Messages vs Single vs Date Range */}
            <div className="flex flex-wrap items-center gap-1.5">
              <div className="inline-flex rounded-lg bg-slate-100 p-0.5 text-[11px] font-semibold">
                <button
                  type="button"
                  onClick={() => {
                    setDateMode("all");
                    setDateFrom("");
                    setDateTo("");
                  }}
                  className={`rounded-md px-2.5 py-1 transition ${
                    dateMode === "all"
                      ? "bg-emerald-600 text-white shadow-2xs"
                      : "text-slate-600 hover:text-slate-900"
                  }`}
                >
                  All Messages
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setDateMode("single");
                    setDateTo("");
                  }}
                  className={`rounded-md px-2.5 py-1 transition ${
                    dateMode === "single"
                      ? "bg-white text-slate-900 shadow-2xs"
                      : "text-slate-600 hover:text-slate-900"
                  }`}
                >
                  Single Date
                </button>
                <button
                  type="button"
                  onClick={() => setDateMode("range")}
                  className={`rounded-md px-2.5 py-1 transition ${
                    dateMode === "range"
                      ? "bg-white text-slate-900 shadow-2xs"
                      : "text-slate-600 hover:text-slate-900"
                  }`}
                >
                  Multi Date
                </button>
              </div>

              {dateMode === "single" && (
                <input
                  type="date"
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                  className="rounded-lg border border-slate-300 bg-white px-2.5 py-1 text-xs outline-none focus:border-emerald-500"
                />
              )}

              {dateMode === "range" && (
                <div className="flex items-center gap-1">
                  <input
                    type="date"
                    value={dateFrom}
                    onChange={(e) => setDateFrom(e.target.value)}
                    className="rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs outline-none focus:border-emerald-500"
                  />
                  <span className="text-xs text-slate-400">to</span>
                  <input
                    type="date"
                    value={dateTo}
                    min={dateFrom || undefined}
                    onChange={(e) => setDateTo(e.target.value)}
                    className="rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs outline-none focus:border-emerald-500"
                  />
                </div>
              )}
            </div>

            <div className="flex items-center gap-2">
              <span className="text-[11px] font-semibold text-slate-500">
                {logGroups.length} farmer{logGroups.length === 1 ? "" : "s"}
              </span>
              {hasLogFilter && (
                <button
                  type="button"
                  onClick={resetLogFilters}
                  className="inline-flex items-center gap-1 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 px-2.5 py-1 text-[11px] font-semibold text-slate-700"
                >
                  <RotateCcw size={12} /> Reset
                </button>
              )}
            </div>
          </div>

          {/* Search, User, District, Upazila Grid */}
          <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
            <div className="relative col-span-2 sm:col-span-1">
              <Search
                size={13}
                className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400"
              />
              <input
                value={logSearch}
                onChange={(e) => setLogSearch(e.target.value)}
                placeholder="Farmer, mobile, message..."
                className={`${selectCls} pl-7`}
              />
            </div>

            <div className="relative">
              <User
                size={12}
                className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none"
              />
              <select
                className={`${selectCls} pl-7`}
                value={logUser}
                onChange={(e) => setLogUser(e.target.value)}
              >
                <option value="">All Users</option>
                {trackedUsers.map((u) => (
                  <option key={u} value={u}>
                    {u}
                  </option>
                ))}
              </select>
            </div>

            <select
              className={selectCls}
              value={logDistrict}
              onChange={(e) => {
                setLogDistrict(e.target.value);
                setLogUpazila("");
              }}
            >
              <option value="">All Districts</option>
              {selDistricts.map((x) => (
                <option key={x} value={x}>
                  {x}
                </option>
              ))}
            </select>

            <select
              className={selectCls}
              value={logUpazila}
              onChange={(e) => setLogUpazila(e.target.value)}
            >
              <option value="">All Upazilas</option>
              {logUpazilas.map((x) => (
                <option key={x} value={x}>
                  {x}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Farmer-wise Message Cards (Mobile-first, no horizontal scroll required) */}
        {loading ? (
          <div className="rounded-xl border border-slate-200 bg-white p-8 text-center text-xs text-slate-400">
            Loading tracking messages...
          </div>
        ) : logGroups.length === 0 ? (
          <div className="rounded-xl border border-slate-200 bg-white p-8 text-center text-xs text-slate-400">
            {hasLogFilter
              ? "No tracking messages match your filters."
              : "No tracking messages yet. Tap 'Track a Farmer' above to add one."}
          </div>
        ) : (
          <div className="space-y-2.5">
            {logGroups
              .slice(0, visibleFarmersCount)
              .map(({ farmer, notes }) => (
                <FarmerTrackingGroupCard
                  key={farmer.id}
                  farmer={farmer}
                  notes={notes}
                  onRetrack={() => setModal({ farmer, tab: "track" })}
                  onDetails={() => setModal({ farmer, tab: "details" })}
                  onViewMap={() => openMap(farmer)}
                />
              ))}

            {logGroups.length > visibleFarmersCount && (
              <div className="pt-1 text-center">
                <button
                  type="button"
                  onClick={() => setVisibleFarmersCount((c) => c + 20)}
                  className="rounded-xl border border-slate-300 bg-white hover:bg-slate-50 px-5 py-2 text-xs font-bold text-slate-700 shadow-2xs"
                >
                  Load more farmers ({logGroups.length - visibleFarmersCount}{" "}
                  remaining)
                </button>
              </div>
            )}
          </div>
        )}
      </section>

      {/* -------------------------------- modals ---------------------------- */}
      {showPicker && (
        <MapPickerModal
          farmers={farmers}
          initDistrict={selDistrict}
          initUpazila={selUpazila}
          initAgent={selAgent}
          onPick={pickFarmer}
          onClose={() => setShowPicker(false)}
        />
      )}

      {modal && (
        <FarmerModal
          key={`${modal.farmer.id}-${modal.tab}`}
          farmer={modal.farmer}
          notes={modalNotes}
          initialTab={modal.tab}
          userName={currentUserName}
          onSave={saveNote}
          onViewMap={openMap}
          onClose={() => setModal(null)}
        />
      )}

      {mapFarmer && (
        <FarmerMapModal
          farmer={mapFarmer}
          onClose={() => setMapFarmer(null)}
          onDetails={() => {
            setModal({ farmer: mapFarmer, tab: "details" });
            setMapFarmer(null);
          }}
        />
      )}

      {toast && (
        <div className="fixed bottom-4 right-4 z-[3000] flex items-center gap-2 rounded-lg bg-slate-900 px-3.5 py-2 text-xs font-semibold text-white shadow-lg">
          <CheckCircle2 size={14} className="text-emerald-400" /> {toast}
        </div>
      )}
    </div>
  );
}