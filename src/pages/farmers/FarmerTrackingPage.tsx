import React, { useEffect, useMemo, useState } from "react";
import { supabase } from "../../lib/supabase";
import { getFarmers } from "../../services/farmerService";
import { useAuth } from "../../context/AuthContext";
import { MapContainer, Marker, useMap } from "react-leaflet";
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
  Crosshair,
  History,
  CheckCircle2,
} from "lucide-react";

/* -------------------------------------------------------------------------- */
/*  Helpers (kept local so FarmersPage.tsx stays untouched)                    */
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

function buildMarker(farmer: any, highlight = false) {
  const { emoji, border } = getFarmerBadgeInfo(farmer);
  const name = esc((farmer.farmer_name || "Farmer").slice(0, 16));
  const ring = highlight
    ? "box-shadow:0 0 0 4px rgba(16,185,129,0.35);"
    : "box-shadow:0 1px 2px rgba(0,0,0,0.15);";
  const html = `
    <div style="display:flex;flex-direction:column;align-items:center;transform:translateY(-10px);">
      <span style="font-size:${highlight ? 30 : 24}px;line-height:1;filter:drop-shadow(0 2px 3px rgba(0,0,0,0.35));">${emoji}</span>
      <div style="margin-top:2px;background:rgba(255,255,255,0.96);border:1px solid ${border};color:#0f172a;font-size:10px;font-weight:700;padding:0 5px;border-radius:4px;white-space:nowrap;${ring}">
        ${name}
      </div>
    </div>`;
  return L.divIcon({
    className: "ft-div-icon",
    html,
    iconSize: [60, 48],
    iconAnchor: [30, 38],
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

function FitBounds({ positions }: { positions: [number, number][] }) {
  const map = useMap();
  useEffect(() => {
    if (positions.length === 0) return;
    if (positions.length === 1) {
      map.setView(positions[0], 15);
      return;
    }
    map.fitBounds(positions, { padding: [45, 45] });
  }, [positions, map]);
  return null;
}

const selectCls =
  "w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-xs text-slate-800 outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500";
const labelCls = "block text-[11px] font-medium text-slate-500 mb-0.5";

/* -------------------------------------------------------------------------- */
/*  Small shared pieces                                                        */
/* -------------------------------------------------------------------------- */

function NoteThread({ notes }: { notes: any[] }) {
  if (notes.length === 0) {
    return (
      <div className="rounded-lg border border-dashed border-slate-300 bg-slate-50 p-4 text-center text-xs text-slate-400">
        No tracking messages yet. Write the first one above.
      </div>
    );
  }
  return (
    <div className="space-y-2">
      {notes.map((n, i) => (
        <div
          key={n.id}
          className={`rounded-lg border px-3 py-2 ${
            i === 0
              ? "border-emerald-300 bg-emerald-50"
              : "border-slate-200 bg-white"
          }`}
        >
          <div className="flex items-center justify-between gap-2 text-[11px]">
            <span className="flex items-center gap-1.5 font-semibold text-slate-800">
              <span className="flex h-5 w-5 items-center justify-center rounded-full bg-slate-800 text-[10px] font-bold text-white">
                {(n.user_name || "?").charAt(0).toUpperCase()}
              </span>
              {n.user_name || "Unknown"}
              {i === 0 && (
                <span className="rounded bg-emerald-600 px-1.5 py-px text-[10px] font-semibold text-white">
                  Latest
                </span>
              )}
            </span>
            <span className="text-slate-500">{fmtDateTime(n.created_at)}</span>
          </div>
          <p className="mt-1 whitespace-pre-wrap text-[13px] leading-snug text-slate-900">
            {n.note}
          </p>
        </div>
      ))}
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
    { label: "Feed", value: farmer.feed_type || "-" },
    { label: "Status", value: farmer.status || "Active" },
  ];

  return (
    <div className="space-y-2 text-xs">
      <div className="grid grid-cols-[84px_1fr] gap-x-2 gap-y-1">
        {rows.map((r) => (
          <React.Fragment key={r.label}>
            <span className="text-slate-500">{r.label}</span>
            <span className="text-slate-900">{r.value}</span>
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

/* -------------------------------------------------------------------------- */
/*  Modals                                                                     */
/* -------------------------------------------------------------------------- */

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
          <button
            onClick={onClose}
            className="text-slate-500 hover:text-slate-800"
          >
            <X size={18} />
          </button>
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

        <div className="min-h-0 flex-1">
          <MapContainer
            center={[24.7471, 90.4203]}
            zoom={8}
            style={{ height: "100%", width: "100%", zIndex: 1 }}
          >
            <MapLayers />
            <FitBounds positions={positions} />
            {withLoc.map((f) => (
              <Marker
                key={f.id}
                position={[Number(f.latitude), Number(f.longitude)]}
                icon={buildMarker(f)}
                eventHandlers={{ click: () => onPick(f) }}
              />
            ))}
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
  return (
    <div className="fixed inset-0 z-[2000] flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="flex h-[80vh] w-full max-w-3xl flex-col overflow-hidden rounded-xl border border-slate-300 bg-white shadow-2xl">
        <div className="flex items-center justify-between border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <div>
            <h2 className="text-sm font-bold text-slate-800">
              {farmer.farmer_name}
            </h2>
            <p className="text-[11px] text-slate-500">
              {[farmer.village, farmer.union_name, farmer.upazila, farmer.district]
                .filter(Boolean)
                .join(", ")}
            </p>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={onDetails}
              className="rounded-md border border-slate-300 bg-white px-2.5 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-50"
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
        <div className="min-h-0 flex-1">
          <MapContainer
            center={pos}
            zoom={17}
            style={{ height: "100%", width: "100%", zIndex: 1 }}
          >
            <MapLayers />
            <Marker position={pos} icon={buildMarker(farmer, true)} />
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
    ["Feed type", farmer.feed_type || "-"],
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
                  placeholder="Write a tracking message..."
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
              <NoteThread notes={notes} />
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
/*  Page                                                                       */
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

  // "Track a farmer" card
  const [mode, setMode] = useState<"select" | "map">("select");
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

  // Tracking log filters
  const [dateMode, setDateMode] = useState<"single" | "range">("single");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [logSearch, setLogSearch] = useState("");
  const [logDistrict, setLogDistrict] = useState("");
  const [logUpazila, setLogUpazila] = useState("");
  const [visibleCount, setVisibleCount] = useState(25);

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

    // Auto-select when the user has only one district / upazila
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
    return m; // each list is already latest-first
  }, [trackings]);

  // selector options
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

  // log filter options
  const logUpazilas = useMemo(
    () =>
      uniqSorted(
        farmers
          .filter((f) => !logDistrict || f.district === logDistrict)
          .map((f) => f.upazila)
      ),
    [farmers, logDistrict]
  );

  const logRows = useMemo(() => {
    const q = logSearch.trim().toLowerCase();
    const groups = new Map<
      string,
      { farmer: any; latest: any; count: number }
    >();
    for (const t of trackings) {
      const key = String(t.farmer_id);
      const f = farmerMap.get(key);
      if (!f) continue;
      const ymd = toYmd(t.created_at);
      if (dateMode === "single") {
        if (dateFrom && ymd !== dateFrom) continue;
      } else {
        if (dateFrom && ymd < dateFrom) continue;
        if (dateTo && ymd > dateTo) continue;
      }
      if (logDistrict && f.district !== logDistrict) continue;
      if (logUpazila && f.upazila !== logUpazila) continue;
      if (
        q &&
        !(
          (f.farmer_name || "").toLowerCase().includes(q) ||
          (f.mobile || "").includes(q) ||
          (f.village || "").toLowerCase().includes(q)
        )
      )
        continue;
      const g = groups.get(key);
      if (!g) groups.set(key, { farmer: f, latest: t, count: 1 });
      else g.count += 1;
    }
    return Array.from(groups.values());
  }, [
    trackings,
    farmerMap,
    dateMode,
    dateFrom,
    dateTo,
    logSearch,
    logDistrict,
    logUpazila,
  ]);

  useEffect(() => {
    setVisibleCount(25);
  }, [dateMode, dateFrom, dateTo, logSearch, logDistrict, logUpazila]);

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
    if (ok) setNoteText("");
  }

  function pickFarmer(f: any) {
    setPickedFarmer(f);
    setFarmerQuery(f.farmer_name || "");
    setComboOpen(false);
    setNoteText("");
    setShowPicker(false);
  }

  function clearPicked() {
    setPickedFarmer(null);
    setFarmerQuery("");
    setNoteText("");
  }

  function resetLogFilters() {
    setDateMode("single");
    setDateFrom("");
    setDateTo("");
    setLogSearch("");
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
    dateFrom || dateTo || logSearch || logDistrict || logUpazila;

  /* --------------------------------- render ------------------------------- */

  return (
    <div className="mx-auto max-w-6xl space-y-4 p-3 sm:p-5">
      <div className="flex items-end justify-between gap-3">
        <div>
          <h1 className="text-lg font-bold text-slate-900 sm:text-xl">
            Farmer Tracking
          </h1>
          <p className="text-xs text-slate-500">
            Log a follow-up message for a farmer and see the latest status of
            every farmer in your area.
          </p>
        </div>
        <span className="hidden rounded-full border border-slate-200 bg-white px-3 py-1 text-xs text-slate-600 sm:inline">
          Signed in as <b className="text-slate-800">{currentUserName}</b>
        </span>
      </div>

      {/* ------------------------- 1. TRACK A FARMER ------------------------ */}
      <section className="rounded-xl border border-slate-200 bg-white shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 px-4 py-2.5">
          <h2 className="text-sm font-bold text-slate-800">Track a farmer</h2>
          <div className="inline-flex rounded-lg bg-slate-100 p-0.5 text-xs font-semibold">
            <button
              onClick={() => setMode("select")}
              className={`rounded-md px-3 py-1 ${
                mode === "select"
                  ? "bg-white text-slate-900 shadow-sm"
                  : "text-slate-500"
              }`}
            >
              Select farmer
            </button>
            <button
              onClick={() => {
                setMode("map");
                setShowPicker(true);
              }}
              className={`flex items-center gap-1 rounded-md px-3 py-1 ${
                mode === "map"
                  ? "bg-white text-slate-900 shadow-sm"
                  : "text-slate-500"
              }`}
            >
              <MapPin size={12} /> Pick from map
            </button>
          </div>
        </div>

        <div className="space-y-3 p-4">
          {mode === "select" ? (
            <div className="grid grid-cols-2 gap-2 md:grid-cols-5">
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
              <div className="relative col-span-2 md:col-span-1">
                <label className={labelCls}>
                  Farmer ({farmerChoices.length})
                </label>
                <div className="relative">
                  <Search
                    size={13}
                    className="absolute left-2 top-1/2 -translate-y-1/2 text-slate-400"
                  />
                  <input
                    value={farmerQuery}
                    onChange={(e) => {
                      setFarmerQuery(e.target.value);
                      setComboOpen(true);
                    }}
                    onFocus={() => setComboOpen(true)}
                    onBlur={() => setTimeout(() => setComboOpen(false), 120)}
                    placeholder="Search name, mobile, village"
                    className={`${selectCls} pl-7`}
                  />
                </div>
                {comboOpen && (
                  <div className="absolute z-30 mt-1 max-h-64 w-full min-w-[260px] overflow-y-auto rounded-lg border border-slate-200 bg-white shadow-lg">
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
                    {farmerChoices.length > 40 && (
                      <div className="p-2 text-center text-[11px] text-slate-400">
                        Showing 40 of {farmerChoices.length}. Type to narrow.
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          ) : (
            <div className="flex items-center justify-between gap-3 rounded-lg border border-dashed border-slate-300 bg-slate-50 px-4 py-3">
              <p className="text-xs text-slate-600">
                Open the map, filter by district, upazila or agent, then tap a
                farmer.
              </p>
              <button
                onClick={() => setShowPicker(true)}
                className="flex shrink-0 items-center gap-1.5 rounded-md bg-slate-800 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-900"
              >
                <Crosshair size={13} /> Open map
              </button>
            </div>
          )}

          {pickedFarmer && (
            <div className="grid gap-3 md:grid-cols-5">
              <div className="rounded-lg border border-slate-200 bg-slate-50 p-3 md:col-span-3">
                <div className="mb-2 flex items-start justify-between gap-2">
                  <div className="flex min-w-0 items-center gap-2.5">
                    <span className="text-3xl leading-none">
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
                  <div className="flex shrink-0 items-center gap-1.5">
                    {pickedFarmer.companies?.logo_url && (
                      <img
                        src={pickedFarmer.companies.logo_url}
                        alt=""
                        className="h-6 max-w-[64px] object-contain"
                      />
                    )}
                    <button
                      onClick={clearPicked}
                      title="Clear"
                      className="rounded p-1 text-slate-400 hover:bg-slate-200 hover:text-slate-700"
                    >
                      <X size={14} />
                    </button>
                  </div>
                </div>
                <FarmerInfoGrid farmer={pickedFarmer} />
                {lastPicked && (
                  <div className="mt-2 rounded-md border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-[11px] text-emerald-900">
                    <span className="font-semibold">Last tracking:</span>{" "}
                    {lastPicked.note.length > 90
                      ? lastPicked.note.slice(0, 90) + "..."
                      : lastPicked.note}{" "}
                    <span className="text-emerald-700">
                      ({lastPicked.user_name}, {timeAgo(lastPicked.created_at)})
                    </span>
                  </div>
                )}
              </div>

              <div className="flex flex-col md:col-span-2">
                <label className={labelCls}>Tracking message</label>
                <textarea
                  value={noteText}
                  onChange={(e) => setNoteText(e.target.value)}
                  placeholder="What did you find? Next step?"
                  className="min-h-[120px] flex-1 resize-none rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500"
                />
                <div className="mt-2 flex items-center justify-between gap-2">
                  <span className="text-[11px] text-slate-500">
                    Tracked by <b className="text-slate-700">{currentUserName}</b>
                  </span>
                  <button
                    onClick={handleSavePicked}
                    disabled={saving || !noteText.trim()}
                    className="flex items-center gap-1.5 rounded-md bg-emerald-600 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700 disabled:opacity-50"
                  >
                    <Send size={12} /> {saving ? "Saving..." : "Save tracking"}
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      </section>

      {/* --------------------------- 2. TRACKING LOG ------------------------- */}
      <section className="rounded-xl border border-slate-200 bg-white shadow-sm">
        <div className="flex items-center justify-between border-b border-slate-100 px-4 py-2.5">
          <h2 className="flex items-center gap-1.5 text-sm font-bold text-slate-800">
            <History size={15} className="text-slate-500" /> Tracking messages
          </h2>
          <span className="text-xs text-slate-500">
            {logRows.length} farmer{logRows.length === 1 ? "" : "s"}
          </span>
        </div>

        <div className="grid grid-cols-2 items-end gap-2 border-b border-slate-100 px-4 py-3 md:grid-cols-6">
          <div className="col-span-2 md:col-span-2">
            <label className={labelCls}>Date</label>
            <div className="flex items-center gap-1.5">
              <div className="inline-flex shrink-0 rounded-md bg-slate-100 p-0.5 text-[11px] font-semibold">
                <button
                  onClick={() => {
                    setDateMode("single");
                    setDateTo("");
                  }}
                  className={`rounded px-2 py-1 ${
                    dateMode === "single"
                      ? "bg-white text-slate-900 shadow-sm"
                      : "text-slate-500"
                  }`}
                >
                  Single
                </button>
                <button
                  onClick={() => setDateMode("range")}
                  className={`rounded px-2 py-1 ${
                    dateMode === "range"
                      ? "bg-white text-slate-900 shadow-sm"
                      : "text-slate-500"
                  }`}
                >
                  Range
                </button>
              </div>
              <input
                type="date"
                value={dateFrom}
                onChange={(e) => setDateFrom(e.target.value)}
                className={selectCls}
              />
              {dateMode === "range" && (
                <>
                  <span className="text-xs text-slate-400">to</span>
                  <input
                    type="date"
                    value={dateTo}
                    min={dateFrom || undefined}
                    onChange={(e) => setDateTo(e.target.value)}
                    className={selectCls}
                  />
                </>
              )}
            </div>
          </div>
          <div className="col-span-2 md:col-span-2">
            <label className={labelCls}>Farmer</label>
            <div className="relative">
              <Search
                size={13}
                className="absolute left-2 top-1/2 -translate-y-1/2 text-slate-400"
              />
              <input
                value={logSearch}
                onChange={(e) => setLogSearch(e.target.value)}
                placeholder="Search name, mobile, village"
                className={`${selectCls} pl-7`}
              />
            </div>
          </div>
          <div>
            <label className={labelCls}>District</label>
            <select
              className={selectCls}
              value={logDistrict}
              onChange={(e) => {
                setLogDistrict(e.target.value);
                setLogUpazila("");
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
          <div className="flex items-end gap-1.5">
            <div className="min-w-0 flex-1">
              <label className={labelCls}>Upazila</label>
              <select
                className={selectCls}
                value={logUpazila}
                onChange={(e) => setLogUpazila(e.target.value)}
              >
                <option value="">All</option>
                {logUpazilas.map((x) => (
                  <option key={x} value={x}>
                    {x}
                  </option>
                ))}
              </select>
            </div>
            {hasLogFilter && (
              <button
                onClick={resetLogFilters}
                title="Reset filters"
                className="rounded-md border border-slate-300 p-1.5 text-slate-600 hover:bg-slate-50"
              >
                <RotateCcw size={14} />
              </button>
            )}
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full min-w-[760px] border-collapse text-xs">
            <thead className="bg-slate-50 text-left text-slate-600">
              <tr>
                <th className="w-[24%] px-4 py-2 font-semibold">Farmer</th>
                <th className="w-[44%] px-3 py-2 font-semibold">
                  Tracking message
                </th>
                <th className="w-[16%] px-3 py-2 font-semibold">Tracked by</th>
                <th className="w-[16%] px-3 py-2 text-right font-semibold">
                  Actions
                </th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={4} className="p-8 text-center text-slate-400">
                    Loading...
                  </td>
                </tr>
              ) : logRows.length === 0 ? (
                <tr>
                  <td colSpan={4} className="p-8 text-center text-slate-400">
                    {hasLogFilter
                      ? "No tracking messages match these filters."
                      : "No tracking messages yet. Track a farmer above to get started."}
                  </td>
                </tr>
              ) : (
                logRows.slice(0, visibleCount).map(({ farmer, latest, count }) => {
                  const badge = getFarmerBadgeInfo(farmer);
                  return (
                    <tr
                      key={farmer.id}
                      className="border-t border-slate-100 align-top hover:bg-slate-50/60"
                    >
                      <td className="px-4 py-3">
                        <button
                          onClick={() =>
                            setModal({ farmer, tab: "details" })
                          }
                          className="flex items-center gap-1.5 text-left"
                        >
                          <span className="text-lg leading-none">
                            {badge.emoji}
                          </span>
                          <span className="text-[13px] font-bold text-slate-900 hover:text-emerald-700 hover:underline">
                            {farmer.farmer_name}
                          </span>
                        </button>
                        <div className="mt-0.5 pl-7 text-[11px] text-slate-500">
                          {farmer.mobile || "No mobile"}
                        </div>
                        <div className="mt-1 flex flex-wrap gap-1 pl-7">
                          {farmer.companies?.company_name && (
                            <span className="rounded bg-slate-100 px-1.5 py-px text-[10px] font-medium text-slate-700">
                              {farmer.companies.company_name}
                            </span>
                          )}
                          {farmer.upazila && (
                            <span className="rounded bg-slate-100 px-1.5 py-px text-[10px] font-medium text-slate-700">
                              {farmer.upazila}
                            </span>
                          )}
                          <span className="rounded bg-amber-50 px-1.5 py-px text-[10px] font-medium text-amber-800">
                            {badge.label}
                          </span>
                        </div>
                      </td>

                      <td className="px-3 py-3">
                        <div className="rounded-lg border-l-4 border-emerald-500 bg-emerald-50/60 px-3 py-2">
                          <p
                            className="whitespace-pre-wrap text-[14px] font-semibold leading-snug text-slate-900"
                            style={{
                              display: "-webkit-box",
                              WebkitLineClamp: 3,
                              WebkitBoxOrient: "vertical",
                              overflow: "hidden",
                            }}
                          >
                            {latest.note}
                          </p>
                          {count > 1 && (
                            <button
                              onClick={() =>
                                setModal({ farmer, tab: "track" })
                              }
                              className="mt-1 text-[11px] font-semibold text-emerald-700 hover:underline"
                            >
                              {count} messages in total
                            </button>
                          )}
                        </div>
                      </td>

                      <td className="px-3 py-3">
                        <div className="text-[13px] font-semibold text-slate-800">
                          {latest.user_name || "Unknown"}
                        </div>
                        <div className="text-[11px] text-slate-500">
                          {fmtDateTime(latest.created_at)}
                        </div>
                        <div className="text-[11px] font-medium text-emerald-700">
                          {timeAgo(latest.created_at)}
                        </div>
                      </td>

                      <td className="px-3 py-3">
                        <div className="flex flex-col items-end gap-1.5">
                          <button
                            onClick={() => setModal({ farmer, tab: "track" })}
                            className="flex w-24 items-center justify-center gap-1 rounded-md bg-emerald-600 px-2.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700"
                          >
                            <RotateCcw size={12} /> Retrack
                          </button>
                          <button
                            onClick={() => openMap(farmer)}
                            disabled={!farmer.latitude || !farmer.longitude}
                            className="flex w-24 items-center justify-center gap-1 rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 disabled:opacity-40"
                          >
                            <MapPin size={12} /> View map
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {logRows.length > visibleCount && (
          <div className="border-t border-slate-100 p-3 text-center">
            <button
              onClick={() => setVisibleCount((c) => c + 25)}
              className="rounded-md border border-slate-300 px-4 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
            >
              Show more ({logRows.length - visibleCount} left)
            </button>
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
