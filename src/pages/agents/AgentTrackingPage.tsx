import React, { useEffect, useMemo, useRef, useState } from "react";
import { supabase } from "../../lib/supabase";
import { useAuth } from "../../context/AuthContext";
import {
  MapContainer,
  Marker,
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
  User,
  Crosshair,
  AlertTriangle,
  Sprout,
  UserCheck,
  Pencil,
  Check,
  AtSign,
} from "lucide-react";
import { useSearchParams } from "react-router-dom";
import { Geolocation } from "@capacitor/geolocation";
import { Capacitor } from "@capacitor/core";
import { UserAvatar, renderNoteWithMentions } from "../../utils/avatarHelper";

/* -------------------------------------------------------------------------- */
/*  Helpers & Status Tag Parser                                               */
/* -------------------------------------------------------------------------- */

const RISKY_PREFIX = "[RISKY]";
const WEAK_PREFIX = "[WEAK_TO_NOURISH]";
const RESOLVED_PREFIX = "[RESOLVED]";

function isNourishAgent(agent: any): boolean {
  const compName = (
    agent?.companies?.company_name ||
    agent?.company_name ||
    ""
  ).toString().toLowerCase();
  return compName.includes("nourish");
}

function parseNoteMeta(rawNote?: string) {
  const text = String(rawNote || "").trim();
  if (text.startsWith(RISKY_PREFIX)) {
    return { type: "risky" as const, badgeText: "Marked Risky", cleanText: text.slice(RISKY_PREFIX.length).trim() };
  }
  if (text.startsWith(WEAK_PREFIX)) {
    return { type: "weak" as const, badgeText: "Weak to Nourish", cleanText: text.slice(WEAK_PREFIX.length).trim() };
  }
  if (text.startsWith(RESOLVED_PREFIX)) {
    return { type: "resolved" as const, badgeText: "Mark Cleared", cleanText: text.slice(RESOLVED_PREFIX.length).trim() };
  }
  return { type: "normal" as const, badgeText: "", cleanText: text };
}

function getAgentSpecialStatus(notes: any[], agent: any) {
  let isRisky = false;
  let riskyNote: any = null;
  let isWeak = false;
  let weakNote: any = null;

  for (const n of notes) {
    const meta = parseNoteMeta(n.note);
    if (meta.type === "resolved") break;
    if (meta.type === "risky" && !riskyNote) {
      isRisky = true;
      riskyNote = { ...n, cleanText: meta.cleanText };
    }
    if (meta.type === "weak" && !weakNote) {
      isWeak = true;
      weakNote = { ...n, cleanText: meta.cleanText };
    }
  }

  const nourish = isNourishAgent(agent);
  return {
    isRisky: nourish && isRisky,
    riskyNote,
    isWeak: !nourish && isWeak,
    weakNote,
  };
}

function parseList(val: any): string[] {
  if (!val) return [];
  if (Array.isArray(val)) return val.map(String).filter(Boolean);
  if (typeof val === "string") {
    try {
      const parsed = JSON.parse(val);
      if (Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
    } catch {
      return val.split(",").map((s) => s.trim()).filter(Boolean);
    }
  }
  return [];
}

function getAgentBadgeInfo(agent: any) {
  return { emoji: "💼", label: agent?.category || "Agent", border: "#2563eb" };
}

function esc(s: string) {
  return (s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function buildMarker(agent: any, highlight = false, distText?: string) {
  const { emoji, border } = getAgentBadgeInfo(agent);
  const name = esc((agent.agent_name || "Agent").slice(0, 16));
  const logo = agent.logo_url || ""; // If you have logo

  const ring = highlight ? "box-shadow:0 0 0 4px rgba(59,130,246,0.35);" : "box-shadow:0 1px 2px rgba(0,0,0,0.15);";
  const distHtml = distText
    ? `<div style="margin-top:2px;background:#0f172a;color:#fff;font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;white-space:nowrap;">${esc(distText)}</div>`
    : "";
    
  const html = `
    <div style="display:flex;flex-direction:column;align-items:center;transform:translateY(-10px);">
      <span style="font-size:${highlight ? 30 : 24}px;line-height:1;filter:drop-shadow(0 2px 3px rgba(0,0,0,0.35));">${logo ? `<img src="${logo}" style="width:30px;height:30px;border-radius:50%;object-fit:cover;background:#fff;"/>` : emoji}</span>
      <div style="margin-top:2px;background:rgba(255,255,255,0.96);border:1px solid ${border};color:#0f172a;font-size:10px;font-weight:700;padding:0 5px;border-radius:4px;white-space:nowrap;${ring}">
        ${name}
      </div>
      ${distHtml}
    </div>`;
    
  return L.divIcon({ className: "ft-div-icon", html, iconSize: [70, 56], iconAnchor: [35, 42] });
}

function toYmd(iso: string) {
  const d = new Date(iso);
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${m}-${day}`;
}

function fmtDateTime(iso: string) {
  return new Date(iso).toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "numeric", minute: "2-digit", hour12: true });
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
  return Array.from(new Set(list.filter(Boolean))).sort((a, b) => a.localeCompare(b));
}

function calcDistanceKm(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const R = 6371;
  const dLat = ((lat2 - lat1) * Math.PI) / 180;
  const dLon = ((lon2 - lon1) * Math.PI) / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos((lat1 * Math.PI) / 180) * Math.cos((lat2 * Math.PI) / 180) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function formatDistance(km: number): string {
  if (km < 1) return `${Math.round(km * 1000)} m`;
  return `${km.toFixed(2)} km`;
}

function FlyToLocation({ target }: { target: [number, number] | null }) {
  const map = useMap();
  useEffect(() => {
    if (target) map.flyTo(target, 15, { animate: true, duration: 0.8 });
  }, [target, map]);
  return null;
}

function FitBounds({ positions, flyTarget }: { positions: [number, number][]; flyTarget?: { pos: [number, number]; zoom?: number; ts: number } | null }) {
  const map = useMap();
  useEffect(() => {
    if (flyTarget) map.flyTo(flyTarget.pos, flyTarget.zoom || 15, { duration: 0.8 });
  }, [flyTarget, map]);
  useEffect(() => {
    if (flyTarget) return;
    if (positions.length === 0) return;
    if (positions.length === 1) { map.setView(positions[0], 15); return; }
    map.fitBounds(positions, { padding: [45, 45] });
  }, [positions, map, flyTarget]);
  return null;
}

/* -------------------------------------------------------------------------- */
/*  Notification Helper                                                       */
/* -------------------------------------------------------------------------- */
async function sendFirebaseNotification(title: string, body: string, targetId: string) {
  try {
    console.log("FCM Triggered:", { title, body, targetId });
  } catch (error) {
    console.error("Failed to send notification", error);
  }
}

const selectCls = "w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs text-slate-800 outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500";
const labelCls = "block text-[11px] font-semibold text-slate-500 mb-0.5";

/* -------------------------------------------------------------------------- */
/*  Shared Components                                                          */
/* -------------------------------------------------------------------------- */

function NoteThread({ notes, initialLimit = 10, currentUserId, currentUserName, isFullAccess, onEditNote, getUserAvatar, onTagUser }: any) {
  const [limit, setLimit] = useState(initialLimit);
  const [editingId, setEditingId] = useState<any>(null);
  const [editValue, setEditValue] = useState("");
  const [savingEdit, setSavingEdit] = useState(false);

  if (notes.length === 0) return <div className="rounded-lg border border-dashed border-slate-300 bg-slate-50 p-4 text-center text-xs text-slate-400">No tracking messages yet.</div>;

  const visibleNotes = notes.slice(0, limit);
  const remaining = notes.length - visibleNotes.length;

  function canEdit(n: any) {
    if (!onEditNote) return false;
    if (isFullAccess) return true;
    if (n.user_id && currentUserId && String(n.user_id) === String(currentUserId)) return true;
    if (n.user_name && currentUserName && n.user_name === currentUserName) return true;
    return false;
  }

  async function handleSaveEdit(n: any) {
    if (!onEditNote || !editValue.trim() || savingEdit) return;
    setSavingEdit(true);
    const ok = await onEditNote(n, editValue.trim());
    setSavingEdit(false);
    if (ok) { setEditingId(null); setEditValue(""); }
  }

  return (
    <div className="space-y-2">
      {visibleNotes.map((n: any, i: number) => {
        const meta = parseNoteMeta(n.note);
        const isRisky = meta.type === "risky";
        const isWeak = meta.type === "weak";
        const isResolved = meta.type === "resolved";
        const isEditing = editingId === n.id;

        const boxStyle = isRisky ? "border-red-300 bg-red-50/70" : isWeak ? "border-teal-300 bg-teal-50/70" : isResolved ? "border-blue-300 bg-blue-50/70" : i === 0 ? "border-blue-400 bg-blue-50/80 shadow-2xs" : "border-slate-200 bg-white";
        const avatarUrl = getUserAvatar ? getUserAvatar(n.user_id, n.user_name) : null;

        return (
          <div key={n.id} id={`track-note-${n.id}`} className={`rounded-lg border px-3 py-2 transition ${boxStyle}`}>
            <div className="flex flex-wrap items-center justify-between gap-1.5 text-[11px]">
              <span className="flex items-center gap-1.5 font-bold text-slate-800 flex-wrap">
                <UserAvatar
                  name={n.user_name || "Officer"}
                  avatarUrl={avatarUrl}
                  size="w-5 h-5"
                  textSize="text-[9px]"
                />
                <span>{n.user_name || "Officer"}</span>
                {isRisky && <span className="inline-flex items-center gap-0.5 rounded bg-red-600 px-1.5 py-px text-[9px] font-bold text-white uppercase tracking-wide">⚠️ Marked Risky</span>}
                {isWeak && <span className="inline-flex items-center gap-0.5 rounded bg-teal-600 px-1.5 py-px text-[9px] font-bold text-white uppercase tracking-wide">🌱 Weak to Nourish</span>}
                {isResolved && <span className="inline-flex items-center gap-0.5 rounded bg-blue-600 px-1.5 py-px text-[9px] font-bold text-white uppercase tracking-wide">✓ Mark Cleared</span>}
                {i === 0 && meta.type === "normal" && <span className="rounded bg-blue-600 px-1.5 py-px text-[9px] font-bold text-white uppercase tracking-wide">Latest</span>}
              </span>
              <div className="flex items-center gap-1.5">
                {onTagUser && (
                  <button
                    type="button"
                    onClick={() => onTagUser(n.user_name || "Officer")}
                    title={`Tag ${n.user_name || "Officer"}`}
                    className="inline-flex items-center gap-0.5 rounded bg-blue-50 hover:bg-blue-100 border border-blue-200 px-1.5 py-0.5 text-[10px] font-bold text-blue-700 transition"
                  >
                    @ Tag
                  </button>
                )}
                <span className="text-slate-500 font-medium">{fmtDateTime(n.created_at)} ({timeAgo(n.created_at)})</span>
                {canEdit(n) && !isEditing && (
                  <button type="button" onClick={() => { setEditingId(n.id); setEditValue(meta.cleanText); }} className="rounded p-1 text-slate-400 hover:bg-slate-200/70 hover:text-blue-700 transition">
                    <Pencil size={11} />
                  </button>
                )}
              </div>
            </div>
            {isEditing ? (
              <div className="mt-1.5 space-y-1.5">
                <textarea rows={2} autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)} className="w-full resize-none rounded-lg border border-blue-400 bg-white px-2.5 py-1.5 text-xs text-slate-900 outline-none focus:ring-1 focus:ring-blue-500" />
                <div className="flex items-center justify-end gap-1.5">
                  <button type="button" onClick={() => { setEditingId(null); setEditValue(""); }} className="rounded-md border border-slate-300 bg-white px-2.5 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-50">Cancel</button>
                  <button type="button" disabled={savingEdit || !editValue.trim()} onClick={() => handleSaveEdit(n)} className="inline-flex items-center gap-1 rounded-md bg-blue-600 hover:bg-blue-700 px-2.5 py-1 text-[11px] font-bold text-white disabled:opacity-50">
                    <Check size={11} />{savingEdit ? "Saving..." : "Update"}
                  </button>
                </div>
              </div>
            ) : (
              <p className="mt-1 whitespace-pre-wrap break-words text-[13px] sm:text-sm font-medium leading-snug text-slate-900">
                {renderNoteWithMentions(meta.cleanText)}
              </p>
            )}
          </div>
        );
      })}
      {remaining > 0 && (
        <button type="button" onClick={() => setLimit((prev: number) => prev + 10)} className="w-full rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 py-1.5 text-center text-xs font-semibold text-blue-700 transition">
          Load more messages ({remaining} older)
        </button>
      )}
    </div>
  );
}

function AgentInfoGrid({ agent }: { agent: any }) {
  const address = [agent.village, agent.union_name, agent.upazila, agent.district].filter(Boolean).join(", ");
  const rows = [
    { label: "Company", value: agent.companies?.company_name || agent.company_name || "-" },
    { label: "Mobile", value: agent.mobile ? <a href={`tel:${agent.mobile}`} className="inline-flex items-center gap-1 font-semibold text-blue-700 hover:underline"><Phone size={11} /> {agent.mobile}</a> : "-" },
    { label: "Address", value: address || "-" },
    { label: "Category", value: agent.category || "Agent" },
    { label: "Status", value: agent.status || "Active" },
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
    </div>
  );
}

function MapPickerModal({ agents, initDistrict, initUpazila, onPick, onClose }: any) {
  const [d, setD] = useState(initDistrict);
  const [u, setU] = useState(initUpazila);
  const [myLoc, setMyLoc] = useState<[number, number] | null>(null);
  const [flyTarget, setFlyTarget] = useState<[number, number] | null>(null);
  const [locating, setLocating] = useState(false);

  const districts = useMemo(() => uniqSorted(agents.map((a: any) => a.district)), [agents]);
  const upazilas = useMemo(() => uniqSorted(agents.filter((a: any) => !d || a.district === d).map((a: any) => a.upazila)), [agents, d]);
  const shown = useMemo(() => agents.filter((a: any) => (!d || a.district === d) && (!u || a.upazila === u)), [agents, d, u]);
  const withLoc = useMemo(() => shown.filter((a: any) => a.latitude && a.longitude), [shown]);
  const positions = useMemo(() => withLoc.map((a: any) => [Number(a.latitude), Number(a.longitude)] as [number, number]), [withLoc]);

  async function handleLocateMe() {
    setLocating(true);
    try {
      if (Capacitor.isNativePlatform()) {
        let perm = await Geolocation.checkPermissions();
        if (perm.location !== "granted") perm = await Geolocation.requestPermissions();
        if (perm.location !== "granted") { alert("Location permission denied"); setLocating(false); return; }
        const pos = await Geolocation.getCurrentPosition({ enableHighAccuracy: true });
        setMyLoc([pos.coords.latitude, pos.coords.longitude]);
        setFlyTarget([pos.coords.latitude, pos.coords.longitude]);
      } else {
        navigator.geolocation.getCurrentPosition(
          (pos) => { setMyLoc([pos.coords.latitude, pos.coords.longitude]); setFlyTarget([pos.coords.latitude, pos.coords.longitude]); },
          () => alert("Location denied in browser.")
        );
      }
    } catch (err: any) { alert("Error: " + err.message); } finally { setLocating(false); }
  }

  return (
    <div className="fixed inset-0 z-[2000] flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="flex h-[88vh] w-full max-w-5xl flex-col overflow-hidden rounded-xl border border-slate-300 bg-white shadow-2xl">
        <div className="flex items-center justify-between border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <div>
            <h2 className="text-sm font-bold text-slate-800">Pick an agent from the map</h2>
            <p className="text-[11px] text-slate-500">{withLoc.length} on map. Tap an agent to select.</p>
          </div>
          <div className="flex items-center gap-2">
            <button type="button" onClick={handleLocateMe} disabled={locating} className="flex items-center gap-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 px-3 py-1.5 text-xs font-bold text-white shadow-sm disabled:opacity-60">
              <Crosshair size={13} /> {locating ? "Locating..." : "Locate Me"}
            </button>
            <button onClick={onClose} className="text-slate-500 hover:text-slate-800"><X size={18} /></button>
          </div>
        </div>
        <div className="grid grid-cols-2 gap-2 border-b border-slate-200 px-4 py-2">
          <div>
            <label className={labelCls}>District</label>
            <select className={selectCls} value={d} onChange={(e) => { setD(e.target.value); setU(""); }}>
              <option value="">All</option>
              {districts.map(x => <option key={x} value={x}>{x}</option>)}
            </select>
          </div>
          <div>
            <label className={labelCls}>Upazila</label>
            <select className={selectCls} value={u} onChange={(e) => setU(e.target.value)}>
              <option value="">All</option>
              {upazilas.map(x => <option key={x} value={x}>{x}</option>)}
            </select>
          </div>
        </div>
        <div className="relative min-h-0 flex-1">
          <MapContainer center={[24.7471, 90.4203]} zoom={8} style={{ height: "100%", width: "100%", zIndex: 1 }}>
            <MapLayers />
            <FitBounds positions={positions} />
            <FlyToLocation target={flyTarget} />
            {myLoc && <CircleMarker center={myLoc} radius={9} pathOptions={{ color: "#ffffff", weight: 3, fillColor: "#2563eb", fillOpacity: 1 }} />}
            {withLoc.map((a: any) => {
              const distText = myLoc ? formatDistance(calcDistanceKm(myLoc[0], myLoc[1], Number(a.latitude), Number(a.longitude))) : undefined;
              return <Marker key={a.id} position={[Number(a.latitude), Number(a.longitude)]} icon={buildMarker(a, false, distText)} eventHandlers={{ click: () => onPick(a) }} />;
            })}
          </MapContainer>
        </div>
      </div>
    </div>
  );
}

function AgentMapModal({ agent, onClose, onDetails }: any) {
  const pos: [number, number] = [Number(agent.latitude), Number(agent.longitude)];
  const [myLoc, _setMyLoc] = useState<[number, number] | null>(null);
  
  const distanceKm = useMemo(() => myLoc ? calcDistanceKm(myLoc[0], myLoc[1], pos[0], pos[1]) : null, [myLoc, pos]);
  const boundsPositions = useMemo<[number, number][]>(() => myLoc ? [myLoc, pos] : [pos], [myLoc, pos]);

  return (
    <div className="fixed inset-0 z-[2000] flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="flex h-[80vh] w-full max-w-3xl flex-col overflow-hidden rounded-xl border border-slate-300 bg-white shadow-2xl">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <div>
            <h2 className="text-sm font-bold text-slate-800">{agent.agent_name}</h2>
            {distanceKm !== null && <span className="mt-1 inline-block rounded-md bg-blue-100 border border-blue-300 px-2 py-0.5 text-[11px] font-bold text-blue-900">Distance from you: {formatDistance(distanceKm)}</span>}
          </div>
          <div className="flex items-center gap-2">
            <button onClick={onDetails} className="rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50">View details</button>
            <button onClick={onClose} className="text-slate-500 hover:text-slate-800"><X size={18} /></button>
          </div>
        </div>
        <div className="relative min-h-0 flex-1">
          <MapContainer center={pos} zoom={16} style={{ height: "100%", width: "100%", zIndex: 1 }}>
            <MapLayers />
            <FitBounds positions={boundsPositions} />
            <Marker position={pos} icon={buildMarker(agent, true, distanceKm !== null ? `${formatDistance(distanceKm)} away` : undefined)} />
          </MapContainer>
        </div>
      </div>
    </div>
  );
}

function AgentModal({ agent, notes, initialTab, initialTagUser, userId, userName, isFullAccess, allUsers, getUserAvatar, onSave, onEditNote, onViewMap, onClose }: any) {
  const [tab, setTab] = useState<"track" | "details">(initialTab);
  const [text, setText] = useState(() => initialTagUser ? `@[${initialTagUser}] ` : "");
  const [postMode, setPostMode] = useState<"normal" | "risky" | "weak" | "clear">("normal");
  const [saving, setSaving] = useState(false);
  const [showMentionList, setShowMentionList] = useState(false);
  const modalTextareaRef = useRef<HTMLTextAreaElement | null>(null);

  const badge = getAgentBadgeInfo(agent);
  const nourish = isNourishAgent(agent);
  const statusInfo = getAgentSpecialStatus(notes, agent);

  function handleModalTagUser(targetName: string) {
    if (!targetName) return;
    const token = `@[${targetName.trim()}] `;
    setText((prev) => (prev.includes(token.trim()) ? prev : `${token}${prev}`));
    setShowMentionList(false);
    setTimeout(() => modalTextareaRef.current?.focus(), 40);
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    if (!text.trim() || saving) return;
    setSaving(true);
    const prefix = postMode === "risky" ? `${RISKY_PREFIX} ` : postMode === "weak" ? `${WEAK_PREFIX} ` : postMode === "clear" ? `${RESOLVED_PREFIX} ` : "";
    const ok = await onSave(agent, `${prefix}${text.trim()}`);
    setSaving(false);
    if (ok) { setText(""); setPostMode("normal"); }
  }

  const tabBtn = (id: "track" | "details", label: string) => (
    <button onClick={() => setTab(id)} className={`border-b-2 px-3 py-2 text-xs font-semibold ${tab === id ? "border-blue-600 text-blue-700" : "border-transparent text-slate-500 hover:text-slate-800"}`}>
      {label}
    </button>
  );

  return (
    <div className="fixed inset-0 z-[2000] flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="flex max-h-[90vh] w-full max-w-xl flex-col overflow-hidden rounded-xl border border-slate-300 bg-white shadow-2xl">
        <div className="flex items-center justify-between gap-3 border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <div className="flex min-w-0 items-center gap-2.5">
            <span className="text-2xl leading-none">{badge.emoji}</span>
            <div className="min-w-0">
              <div className="flex flex-wrap items-center gap-1.5">
                <h2 className="truncate text-sm font-bold text-slate-800">{agent.agent_name}</h2>
                {statusInfo.isRisky && <span className="rounded-full bg-red-600 px-2 py-0.5 text-[10px] font-bold text-white">⚠️ Risky</span>}
                {statusInfo.isWeak && <span className="rounded-full bg-teal-600 px-2 py-0.5 text-[10px] font-bold text-white">🌱 Weak to Nourish</span>}
              </div>
              <p className="truncate text-[11px] text-slate-500">{agent.mobile || "No mobile"}</p>
            </div>
          </div>
          <div className="flex shrink-0 items-center gap-2">
            {agent.latitude && agent.longitude && (
              <button onClick={() => onViewMap(agent)} className="flex items-center gap-1 rounded-md border border-slate-300 bg-white px-2 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-50">
                <MapPin size={12} /> Map
              </button>
            )}
            <button onClick={onClose} className="text-slate-500 hover:text-slate-800"><X size={18} /></button>
          </div>
        </div>

        <div className="flex border-b border-slate-200 px-2">
          {tabBtn("track", `Tracking (${notes.length})`)}
          {tabBtn("details", "Details")}
        </div>

        <div className="overflow-y-auto p-4">
          {tab === "track" ? (
            <div className="space-y-3">
              <div className="relative flex flex-wrap items-center justify-between gap-1.5">
                <div className="flex flex-wrap items-center gap-1.5">
                  {nourish ? (
                    statusInfo.isRisky ? (
                      <button type="button" onClick={() => setPostMode(postMode === "clear" ? "normal" : "clear")} className={`inline-flex items-center gap-1 rounded-lg border px-2.5 py-1 text-[11px] font-bold transition ${postMode === "clear" ? "border-blue-600 bg-blue-600 text-white" : "border-blue-300 bg-blue-50 text-blue-800 hover:bg-blue-100"}`}>
                        <UserCheck size={11} /> Clear Risky Mark
                      </button>
                    ) : (
                      <button type="button" onClick={() => setPostMode(postMode === "risky" ? "normal" : "risky")} className={`inline-flex items-center gap-1 rounded-lg border px-2.5 py-1 text-[11px] font-bold transition ${postMode === "risky" ? "border-red-600 bg-red-600 text-white" : "border-red-200 bg-red-50 text-red-700 hover:bg-red-100"}`}>
                        <AlertTriangle size={11} /> Mark as Risky
                      </button>
                    )
                  ) : statusInfo.isWeak ? (
                    <button type="button" onClick={() => setPostMode(postMode === "clear" ? "normal" : "clear")} className={`inline-flex items-center gap-1 rounded-lg border px-2.5 py-1 text-[11px] font-bold transition ${postMode === "clear" ? "border-blue-600 bg-blue-600 text-white" : "border-blue-300 bg-blue-50 text-blue-800 hover:bg-blue-100"}`}>
                      <UserCheck size={11} /> Clear Weak Mark
                    </button>
                  ) : (
                    <button type="button" onClick={() => setPostMode(postMode === "weak" ? "normal" : "weak")} className={`inline-flex items-center gap-1 rounded-lg border px-2.5 py-1 text-[11px] font-bold transition ${postMode === "weak" ? "border-teal-600 bg-teal-600 text-white" : "border-teal-200 bg-teal-50 text-teal-800 hover:bg-teal-100"}`}>
                      <Sprout size={11} /> Weak to Nourish
                    </button>
                  )}
                </div>

                <button
                  type="button"
                  onClick={() => setShowMentionList((v) => !v)}
                  className="inline-flex items-center gap-1 rounded-lg border border-blue-200 bg-blue-50 hover:bg-blue-100 px-2.5 py-1 text-[11px] font-bold text-blue-700 transition"
                >
                  <AtSign size={11} /> Tag User
                </button>

                {showMentionList && (
                  <div className="absolute right-0 top-8 z-40 w-56 max-h-48 overflow-y-auto rounded-xl border border-slate-200 bg-white p-1 shadow-xl space-y-0.5">
                    {(allUsers || []).map((u: any) => {
                      const uName = u.full_name || u.email?.split("@")[0] || "Officer";
                      return (
                        <button
                          key={u.id}
                          type="button"
                          onClick={() => handleModalTagUser(uName)}
                          className="w-full flex items-center gap-2 px-2 py-1.5 text-left text-xs hover:bg-blue-50 rounded-lg transition"
                        >
                          <UserAvatar name={uName} avatarUrl={u.avatar_url} size="w-5 h-5" textSize="text-[9px]" />
                          <span className="font-semibold text-slate-700 truncate">{uName}</span>
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>

              <form onSubmit={submit} className="space-y-1.5">
                <textarea ref={modalTextareaRef} autoFocus rows={3} value={text} onChange={(e) => setText(e.target.value)}
                  placeholder={postMode === "risky" ? "Write comment for marking as Risky..." : postMode === "weak" ? "Write comment for marking as Weak..." : postMode === "clear" ? "Write comment why clearing mark..." : "Write a new tracking message or @tag someone..."}
                  className={`w-full resize-none rounded-lg border px-3 py-2 text-sm outline-none transition ${postMode === "risky" ? "border-red-400 bg-red-50/40 focus:border-red-600" : postMode === "weak" ? "border-teal-400 bg-teal-50/40 focus:border-teal-600" : postMode === "clear" ? "border-blue-500 bg-blue-50/60 focus:border-blue-700" : "border-slate-300 bg-white focus:border-blue-500"}`}
                />
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-slate-500">Posting as <b className="text-slate-700">{userName}</b></span>
                  <button type="submit" disabled={saving || !text.trim()} className={`flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-semibold text-white disabled:opacity-50 ${postMode === "risky" ? "bg-red-600 hover:bg-red-700" : postMode === "weak" ? "bg-teal-600 hover:bg-teal-700" : "bg-blue-600 hover:bg-blue-700"}`}>
                    <Send size={12} /> {saving ? "Saving..." : postMode === "risky" ? "Save Risky & Notify" : postMode === "weak" ? "Save Weak & Notify" : postMode === "clear" ? "Clear Mark" : "Save tracking"}
                  </button>
                </div>
              </form>
              <NoteThread
                notes={notes}
                initialLimit={10}
                currentUserId={userId}
                currentUserName={userName}
                isFullAccess={isFullAccess}
                onEditNote={onEditNote}
                getUserAvatar={getUserAvatar}
                onTagUser={handleModalTagUser}
              />
            </div>
          ) : (
            <AgentInfoGrid agent={agent} />
          )}
        </div>
      </div>
    </div>
  );
}

function AgentTrackingGroupCard({ agent, notes, currentUserId, currentUserName, isFullAccess, getUserAvatar, onTagFromNote, onRetrack, onDetails, onViewMap, onEditNote }: any) {
  const badge = getAgentBadgeInfo(agent);
  const statusInfo = getAgentSpecialStatus(notes, agent);

  return (
    <div id={`agent-track-card-${agent.id}`} className={`rounded-xl border bg-white p-3 sm:p-4 shadow-2xs transition ${statusInfo.isRisky ? "border-red-300" : statusInfo.isWeak ? "border-teal-300" : "border-slate-200 hover:border-slate-300"}`}>
      <div className="grid grid-cols-1 gap-3 lg:grid-cols-12">
        <div className="flex flex-col justify-between gap-2.5 border-b border-slate-100 pb-2.5 lg:col-span-4 lg:border-b-0 lg:border-r lg:pb-0 lg:pr-3">
          <div className="flex items-start justify-between gap-2">
            <button type="button" onClick={onDetails} className="flex items-start gap-2.5 text-left min-w-0">
              <span className="text-xl leading-none shrink-0 mt-0.5">{badge.emoji}</span>
              <div className="min-w-0">
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="truncate text-sm font-bold text-slate-900 hover:text-blue-700 hover:underline">{agent.agent_name}</span>
                  {statusInfo.isRisky && <span className="inline-flex items-center gap-0.5 rounded-full bg-red-100 border border-red-200 px-1.5 py-px text-[9px] font-bold text-red-700">⚠️ Risky</span>}
                  {statusInfo.isWeak && <span className="inline-flex items-center gap-0.5 rounded-full bg-teal-100 border border-teal-200 px-1.5 py-px text-[9px] font-bold text-teal-800">🌱 Weak to Nourish</span>}
                </div>
                <div className="truncate text-[11px] font-medium text-slate-500 mt-0.5">{agent.companies?.company_name || agent.company_name || "-"}</div>
              </div>
            </button>
            <span className="shrink-0 rounded-full bg-blue-50 border border-blue-200 px-2 py-0.5 text-[10px] font-bold text-blue-800">{notes.length} msg</span>
          </div>
          <div className="flex items-center justify-between gap-3 pt-1">
            <button type="button" onClick={onRetrack} className="inline-flex items-center justify-center gap-1 rounded-lg bg-blue-600 hover:bg-blue-700 px-2.5 py-1 text-[11px] font-bold text-white shadow-2xs transition"><RotateCcw size={11} /> Retrack</button>
            <button type="button" onClick={onViewMap} disabled={!agent.latitude || !agent.longitude} className="inline-flex items-center justify-center gap-1 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 px-2.5 py-1 text-[11px] font-semibold text-slate-700 disabled:opacity-40 transition"><MapPin size={11} /> Map</button>
          </div>
        </div>
        <div className="lg:col-span-8 min-w-0" id={`agent-notes-${agent.id}`}>
          <NoteThread
            notes={notes}
            initialLimit={10}
            currentUserId={currentUserId}
            currentUserName={currentUserName}
            isFullAccess={isFullAccess}
            onEditNote={onEditNote}
            getUserAvatar={getUserAvatar}
            onTagUser={(targetUserName: string) => onTagFromNote(agent, targetUserName)}
          />
        </div>
      </div>
    </div>
  );
}

/* -------------------------------------------------------------------------- */
/*  Main Page Content                                                          */
/* -------------------------------------------------------------------------- */

export default function AgentTrackingPage() {
  const { role } = useAuth();
  if (role === "viewer") return <div className="p-6 text-center text-slate-500">Access Denied</div>;
  return <TrackingContent role={role} />;
}

function TrackingContent({ role }: { role: any }) {
  const isFullAccess = role === "super_admin" || role === "manager";
  const [currentUserId, setCurrentUserId] = useState("");
  const [currentUserName, setCurrentUserName] = useState("User");
  const [allUsers, setAllUsers] = useState<any[]>([]);
  
  const [agents, setAgents] = useState<any[]>([]);
  const [trackings, setTrackings] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState("");

  const [trackOpen, setTrackOpen] = useState(false);
  const [selDistrict, setSelDistrict] = useState("");
  const [selUpazila, setSelUpazila] = useState("");
  const [selCompany, setSelCompany] = useState("");
  const [agentQuery, setAgentQuery] = useState("");
  const [comboOpen, setComboOpen] = useState(false);
  const [pickedAgent, setPickedAgent] = useState<any>(null);
  const [noteText, setNoteText] = useState("");
  const [pickedMode, setPickedMode] = useState<"normal" | "risky" | "weak" | "clear">("normal");
  const [saving, setSaving] = useState(false);
  const [showPicker, setShowPicker] = useState(false);
  const [showMentionList, setShowMentionList] = useState(false);
  const pickedTextareaRef = useRef<HTMLTextAreaElement | null>(null);

  const [statusTab, setStatusTab] = useState<"all" | "risky" | "weak">("all");
  const [dateMode, setDateMode] = useState<"all" | "single" | "range">("all");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [logSearch, setLogSearch] = useState("");
  const [logUser, setLogUser] = useState("");
  const [logDistrict, setLogDistrict] = useState("");
  const [logUpazila, setLogUpazila] = useState("");
  const [visibleCount, setVisibleCount] = useState(20);

  const [modal, setModal] = useState<{ agent: any; tab: "track" | "details"; tagUser?: string } | null>(null);
  const [mapAgent, setMapAgent] = useState<any>(null);

  /* Data Loading */
  async function loadTrackings() {
    const { data, error } = await supabase.from("agent_trackings").select("*").order("created_at", { ascending: false });
    if (!error && data) setTrackings(data);
  }

  async function loadData() {
    setLoading(true);
    const { data: { user } } = await supabase.auth.getUser();
    let uName = "Officer";
    let uDistricts: string[] = [];
    let uUpazilas: string[] = [];

    if (user) {
      setCurrentUserId(user.id);
      const { data: profile } = await supabase.from("users").select("full_name, districts, upazilas").eq("id", user.id).single();
      if (profile) {
        uName = profile.full_name || "Officer";
        uDistricts = parseList(profile.districts);
        uUpazilas = parseList(profile.upazilas);
      }
    }
    setCurrentUserName(uName);

    const [{ data: agentData }, { data: usersList }] = await Promise.all([
      supabase.from("agents").select("*, companies(company_name)"),
      supabase.from("users").select("id, full_name, email, avatar_url"),
    ]);

    if (usersList) setAllUsers(usersList);

    let allowedAgents = agentData || [];
    if (!isFullAccess) {
      allowedAgents = allowedAgents.filter((a: any) => {
        const matchUpa = uUpazilas.length === 0 || uUpazilas.includes(a.upazila);
        const matchDist = uDistricts.length === 0 || uDistricts.includes(a.district);
        return matchUpa && matchDist;
      });
    }
    setAgents(allowedAgents);
    await loadTrackings();
    setLoading(false);
  }

  useEffect(() => { loadData(); }, [role]);
  useEffect(() => { if (toast) { const t = setTimeout(() => setToast(""), 2500); return () => clearTimeout(t); } }, [toast]);

  function getUserAvatar(userId?: string, userName?: string): string | null {
    const found = allUsers.find(
      (u) =>
        (userId && String(u.id) === String(userId)) ||
        (userName && String(u.full_name || "").trim().toLowerCase() === String(userName).trim().toLowerCase())
    );
    return found?.avatar_url || null;
  }

  function handlePickedTagUser(targetName: string) {
    if (!targetName) return;
    const token = `@[${targetName.trim()}] `;
    setNoteText((prev) => (prev.includes(token.trim()) ? prev : `${token}${prev}`));
    setShowMentionList(false);
    setTimeout(() => pickedTextareaRef.current?.focus(), 40);
  }

  function handleTagFromGroupCard(agent: any, targetUserName: string) {
    setPickedAgent(agent);
    setAgentQuery(agent.agent_name || "");
    setComboOpen(false);
    setPickedMode("normal");
    setTrackOpen(true);
    const token = `@[${targetUserName.trim()}] `;
    setNoteText((prev) => (prev.includes(token.trim()) ? prev : `${token}${prev}`));
    window.scrollTo({ top: 0, behavior: "smooth" });
    setTimeout(() => pickedTextareaRef.current?.focus(), 180);
  }

  // Derived Data
  const agentMap = useMemo(() => new Map(agents.map(a => [String(a.id), a])), [agents]);
  const notesByAgent = useMemo(() => {
    const m = new Map<string, any[]>();
    trackings.forEach(t => {
      const k = String(t.agent_id);
      if (!m.has(k)) m.set(k, []);
      m.get(k)!.push(t);
    });
    return m;
  }, [trackings]);

  const statusByAgent = useMemo(() => {
    const m = new Map();
    agents.forEach(a => m.set(String(a.id), getAgentSpecialStatus(notesByAgent.get(String(a.id)) || [], a)));
    return m;
  }, [agents, notesByAgent]);

  const riskyTotalCount = useMemo(() => Array.from(statusByAgent.values()).filter(s => s.isRisky).length, [statusByAgent]);
  const weakTotalCount = useMemo(() => Array.from(statusByAgent.values()).filter(s => s.isWeak).length, [statusByAgent]);

  const selDistricts = useMemo(() => uniqSorted(agents.map(a => a.district)), [agents]);
  const selUpazilas = useMemo(() => uniqSorted(agents.filter(a => !selDistrict || a.district === selDistrict).map(a => a.upazila)), [agents, selDistrict]);
  const baseByArea = useMemo(() => agents.filter(a => (!selDistrict || a.district === selDistrict) && (!selUpazila || a.upazila === selUpazila)), [agents, selDistrict, selUpazila]);
  
  const companyOptions = useMemo(() => {
    const m = new Map<string, string>();
    baseByArea.forEach(a => { if (a.company_id) m.set(String(a.company_id), a.companies?.company_name || a.company_name || "Company"); });
    return Array.from(m.entries()).sort((x, y) => x[1].localeCompare(y[1]));
  }, [baseByArea]);

  const agentChoices = useMemo(() => {
    const q = agentQuery.toLowerCase();
    return baseByArea
      .filter(a => !selCompany || String(a.company_id) === selCompany)
      .filter(a => !q || (a.agent_name || "").toLowerCase().includes(q) || (a.mobile || "").includes(q) || (a.village || "").toLowerCase().includes(q))
      .sort((x, y) => (x.agent_name || "").localeCompare(y.agent_name || ""));
  }, [baseByArea, selCompany, agentQuery]);

  const trackedUsers = useMemo(() => uniqSorted(trackings.filter(t => agentMap.has(String(t.agent_id)) && t.user_name).map(t => t.user_name)), [trackings, agentMap]);
  const logUpazilas = useMemo(() => uniqSorted(agents.filter(a => !logDistrict || a.district === logDistrict).map(a => a.upazila)), [agents, logDistrict]);

  const logGroups = useMemo(() => {
    const q = logSearch.toLowerCase();
    const groups = new Map();
    for (const t of trackings) {
      const a = agentMap.get(String(t.agent_id));
      if (!a) continue;
      const st = statusByAgent.get(String(t.agent_id));
      if (statusTab === "risky" && !st?.isRisky) continue;
      if (statusTab === "weak" && !st?.isWeak) continue;

      if (dateMode === "single" && dateFrom && toYmd(t.created_at) !== dateFrom) continue;
      if (dateMode === "range") {
        const ymd = toYmd(t.created_at);
        if (dateFrom && ymd < dateFrom) continue;
        if (dateTo && ymd > dateTo) continue;
      }
      if (logUser && t.user_name !== logUser) continue;
      if (logDistrict && a.district !== logDistrict) continue;
      if (logUpazila && a.upazila !== logUpazila) continue;
      if (q && !((a.agent_name || "").toLowerCase().includes(q) || (a.mobile || "").includes(q) || (t.note || "").toLowerCase().includes(q) || (t.user_name || "").toLowerCase().includes(q))) continue;

      if (!groups.has(String(a.id))) groups.set(String(a.id), { agent: a, notes: [t] });
      else groups.get(String(a.id)).notes.push(t);
    }
    return Array.from(groups.values());
  }, [trackings, agentMap, statusByAgent, statusTab, dateMode, dateFrom, dateTo, logSearch, logUser, logDistrict, logUpazila]);

  useEffect(() => { setVisibleCount(20); }, [statusTab, dateMode, dateFrom, dateTo, logSearch, logUser, logDistrict, logUpazila]);

  // Realtime subscription
  useEffect(() => {
    const ch = supabase.channel("page_live_agent_trackings")
      .on("postgres_changes", { event: "*", schema: "public", table: "agent_trackings" }, () => { loadTrackings(); })
      .subscribe();
    return () => { supabase.removeChannel(ch); };
  }, []);

  const [searchParams, setSearchParams] = useSearchParams();
  useEffect(() => {
    if (loading) return;
    const focusAgentId = searchParams.get("focus");
    const focusNoteId = searchParams.get("note");
    if (!focusAgentId && !focusNoteId) return;

    setTrackOpen(false);
    const idx = logGroups.findIndex((g: any) => String(g.agent.id) === String(focusAgentId));
    if (idx >= visibleCount) setVisibleCount(idx + 5);

    const timer = setTimeout(() => {
      const target = (focusNoteId ? document.getElementById(`track-note-${focusNoteId}`) : null) || (focusAgentId ? document.getElementById(`agent-track-card-${focusAgentId}`) : null);
      if (target) {
        target.scrollIntoView({ behavior: "smooth", block: "center" });
        target.classList.add("ring-2", "ring-blue-500", "bg-amber-50");
        setTimeout(() => target.classList.remove("ring-2", "ring-blue-500", "bg-amber-50"), 3000);
      }
      searchParams.delete("focus"); searchParams.delete("note"); setSearchParams(searchParams, { replace: true });
    }, 200);
    return () => clearTimeout(timer);
  }, [loading, logGroups, searchParams]);

  // Actions
  async function saveNote(agent: any, text: string) {
    if (!text.trim()) return false;
    const { data, error } = await supabase.from("agent_trackings").insert([{ agent_id: agent.id, note: text.trim(), user_id: currentUserId || null, user_name: currentUserName }]).select().single();
    if (error) { alert("Failed to save"); return false; }
    
    await sendFirebaseNotification(`New tracking for ${agent.agent_name}`, text.trim(), String(data.id));

    setTrackings(p => [data, ...p]);
    setToast("Tracking saved!");
    return true;
  }

  async function updateNote(noteObj: any, newText: string) {
    const clean = newText.trim();
    if (!clean) return false;
    const meta = parseNoteMeta(noteObj.note);
    const prefix = meta.type === "risky" ? `${RISKY_PREFIX} ` : meta.type === "weak" ? `${WEAK_PREFIX} ` : meta.type === "resolved" ? `${RESOLVED_PREFIX} ` : "";
    const finalNote = `${prefix}${clean}`;

    const { error } = await supabase.from("agent_trackings").update({ note: finalNote }).eq("id", noteObj.id);
    if (!error) {
      setTrackings(p => p.map(x => x.id === noteObj.id ? { ...x, note: finalNote } : x));
      setToast("Updated");
      return true;
    }
    return false;
  }

  async function handleSavePicked() {
    if (!pickedAgent || !noteText.trim() || saving) return;
    setSaving(true);
    const prefix = pickedMode === "risky" ? `${RISKY_PREFIX} ` : pickedMode === "weak" ? `${WEAK_PREFIX} ` : pickedMode === "clear" ? `${RESOLVED_PREFIX} ` : "";
    const ok = await saveNote(pickedAgent, `${prefix}${noteText.trim()}`);
    setSaving(false);
    if (ok) { setNoteText(""); setPickedMode("normal"); setPickedAgent(null); setAgentQuery(""); setTrackOpen(false); }
  }

  function pickAgent(a: any) { setPickedAgent(a); setAgentQuery(a.agent_name || ""); setComboOpen(false); setNoteText(""); setPickedMode("normal"); setShowPicker(false); setTrackOpen(true); }
  function clearPicked() { setPickedAgent(null); setAgentQuery(""); setNoteText(""); setPickedMode("normal"); }
  function resetLogFilters() { setStatusTab("all"); setDateMode("all"); setDateFrom(""); setDateTo(""); setLogSearch(""); setLogUser(""); setLogDistrict(""); setLogUpazila(""); }
  function openMap(a: any) { if (!a.latitude || !a.longitude) alert("No coordinates saved"); else setMapAgent(a); }

  const pickedNotes = pickedAgent ? notesByAgent.get(String(pickedAgent.id)) || [] : [];
  const pickedStatus = pickedAgent ? getAgentSpecialStatus(pickedNotes, pickedAgent) : { isRisky: false, isWeak: false, riskyNote: null, weakNote: null };
  const lastPicked = pickedNotes[0];
  const pickedBadge = pickedAgent ? getAgentBadgeInfo(pickedAgent) : null;
  const hasLogFilter = statusTab !== "all" || dateMode !== "all" || dateFrom || dateTo || logSearch || logUser || logDistrict || logUpazila;

  return (
    <div className="mx-auto max-w-6xl space-y-3 p-2.5 sm:p-4">
      {/* 1. SMOOTH ANIMATED TRACK AN AGENT SECTION */}
      <section className={`relative z-30 rounded-xl border bg-white transition-all duration-300 ${trackOpen ? "border-blue-500/80 shadow-lg ring-2 ring-blue-500/15" : "border-slate-200 shadow-2xs"}`}>
        <div className="flex flex-wrap items-center justify-between gap-2 p-2.5 sm:px-4">
          <div className="flex flex-wrap items-center gap-2">
            <button type="button" onClick={() => setTrackOpen((o) => !o)} className={`inline-flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-bold shadow-xs transition-all duration-300 ${trackOpen ? "bg-slate-800 text-white" : "bg-blue-600 hover:bg-blue-700 text-white"}`}>
              <Plus size={14} className={`transition-transform duration-300 ${trackOpen ? "rotate-45" : "rotate-0"}`} />
              <span>{trackOpen ? "Close Tracking Form" : "Track an Agent"}</span>
              <ChevronDown size={14} className={`transition-transform duration-300 ${trackOpen ? "rotate-180" : "rotate-0"}`} />
            </button>
            <button type="button" onClick={() => setShowPicker(true)} className="inline-flex items-center gap-1.5 rounded-lg border border-blue-200 bg-blue-50 hover:bg-blue-100 px-3 py-2 text-xs font-bold text-blue-700 transition">
              <MapPin size={13} /> Pick from Map
            </button>
          </div>
          {pickedAgent && !trackOpen && <button type="button" onClick={() => setTrackOpen(true)} className="text-xs font-semibold text-blue-700 hover:underline animate-pulse">Selected: {pickedAgent.agent_name} (Click to open)</button>}
        </div>

        <div className={`grid transition-all duration-300 ease-in-out ${trackOpen ? "grid-rows-[1fr] opacity-100" : "grid-rows-[0fr] opacity-0 pointer-events-none"}`}>
          <div className={comboOpen || showMentionList ? "overflow-visible" : "overflow-hidden"}>
            <div className="border-t border-slate-100 bg-slate-50/40 p-3 sm:p-4 space-y-3 rounded-b-xl">
              <div className="grid grid-cols-2 gap-2 sm:grid-cols-3 md:grid-cols-4">
                <div>
                  <label className={labelCls}>District</label>
                  <select className={selectCls} value={selDistrict} onChange={(e) => { setSelDistrict(e.target.value); setSelUpazila(""); setSelCompany(""); }}>
                    <option value="">All</option>
                    {selDistricts.map(x => <option key={x} value={x}>{x}</option>)}
                  </select>
                </div>
                <div>
                  <label className={labelCls}>Upazila</label>
                  <select className={selectCls} value={selUpazila} onChange={(e) => { setSelUpazila(e.target.value); setSelCompany(""); }}>
                    <option value="">All</option>
                    {selUpazilas.map(x => <option key={x} value={x}>{x}</option>)}
                  </select>
                </div>
                <div>
                  <label className={labelCls}>Company</label>
                  <select className={selectCls} value={selCompany} onChange={(e) => setSelCompany(e.target.value)}>
                    <option value="">All</option>
                    {companyOptions.map(([id, name]) => <option key={id} value={id}>{name}</option>)}
                  </select>
                </div>
                <div className="relative col-span-2 sm:col-span-3 md:col-span-1">
                  <label className={labelCls}>Agent ({agentChoices.length})</label>
                  <div className="relative">
                    <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input value={agentQuery} onChange={e => { setAgentQuery(e.target.value); setComboOpen(true); }} onFocus={() => setComboOpen(true)} onBlur={() => setTimeout(() => setComboOpen(false), 180)} placeholder="Search agent..." className={`${selectCls} pl-7`} />
                  </div>
                  {comboOpen && (
                    <div className="absolute right-0 left-0 top-full z-[100] mt-1 max-h-64 overflow-y-auto rounded-xl border border-slate-300 bg-white shadow-2xl">
                      {agentChoices.length === 0 ? <div className="p-3 text-center text-xs text-slate-400">No agent matched.</div> : agentChoices.slice(0, 40).map(a => (
                        <button key={a.id} type="button" onMouseDown={(e) => { e.preventDefault(); pickAgent(a); }} className="flex w-full items-center gap-2 border-b border-slate-100 px-3 py-2 text-left hover:bg-blue-50 transition">
                          <span className="text-lg leading-none">{getAgentBadgeInfo(a).emoji}</span>
                          <span className="min-w-0"><span className="block truncate text-xs font-semibold text-slate-900">{a.agent_name}</span><span className="block truncate text-[11px] text-slate-500">{[a.mobile, a.village, a.companies?.company_name].filter(Boolean).join(", ")}</span></span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {pickedAgent && (
                <div className="grid gap-3 md:grid-cols-5 pt-1 transition-all duration-300">
                  <div className="rounded-lg border border-slate-200 bg-white p-3 md:col-span-3">
                    <div className="mb-2 flex items-start justify-between gap-2">
                      <div className="flex min-w-0 items-center gap-2.5">
                        <span className="text-2xl leading-none">{pickedBadge?.emoji}</span>
                        <div className="min-w-0">
                          <div className="flex flex-wrap items-center gap-1.5">
                            <span className="truncate text-sm font-bold text-slate-900">{pickedAgent.agent_name}</span>
                            {pickedStatus.isRisky && <span className="rounded-full bg-red-600 px-2 py-0.5 text-[10px] font-bold text-white">⚠️ Risky</span>}
                            {pickedStatus.isWeak && <span className="rounded-full bg-teal-600 px-2 py-0.5 text-[10px] font-bold text-white">🌱 Weak to Nourish</span>}
                          </div>
                          <div className="text-[11px] text-slate-500">{pickedBadge?.label} {pickedNotes.length > 0 && `, tracked ${pickedNotes.length} time(s)`}</div>
                        </div>
                      </div>
                      <button onClick={clearPicked} title="Clear" className="rounded p-1 text-slate-400 hover:bg-slate-200 hover:text-slate-700"><X size={14} /></button>
                    </div>
                    <AgentInfoGrid agent={pickedAgent} />
                    {lastPicked && (
                      <div className="mt-2 rounded-md border border-blue-200 bg-blue-50 px-2.5 py-1.5 text-[11px] text-blue-900">
                        <span className="font-semibold">Last tracking:</span> {parseNoteMeta(lastPicked.note).cleanText.length > 90 ? parseNoteMeta(lastPicked.note).cleanText.slice(0, 90) + "..." : parseNoteMeta(lastPicked.note).cleanText} <span className="text-blue-700">({lastPicked.user_name}, {timeAgo(lastPicked.created_at)})</span>
                      </div>
                    )}
                  </div>

                  <div className="flex flex-col md:col-span-2 relative">
                    <div className="mb-1.5 flex flex-wrap items-center justify-between gap-1.5">
                      <label className="text-[11px] font-semibold text-slate-500">Tracking message *</label>
                      <div className="flex flex-wrap items-center gap-1">
                        <button
                          type="button"
                          onClick={() => setShowMentionList((v) => !v)}
                          className="inline-flex items-center gap-0.5 rounded border border-blue-200 bg-blue-50 hover:bg-blue-100 px-2 py-0.5 text-[10px] font-bold text-blue-700 transition"
                        >
                          <AtSign size={10} /> Tag User
                        </button>
                        {isNourishAgent(pickedAgent) ? (
                          pickedStatus.isRisky ? (
                            <button type="button" onClick={() => setPickedMode(pickedMode === "clear" ? "normal" : "clear")} className={`inline-flex items-center gap-1 rounded px-2 py-0.5 text-[10px] font-bold transition ${pickedMode === "clear" ? "bg-blue-600 text-white" : "border border-blue-300 bg-blue-50 text-blue-800 hover:bg-blue-100"}`}><UserCheck size={10} /> Clear Risky Mark</button>
                          ) : (
                            <button type="button" onClick={() => setPickedMode(pickedMode === "risky" ? "normal" : "risky")} className={`inline-flex items-center gap-1 rounded px-2 py-0.5 text-[10px] font-bold transition ${pickedMode === "risky" ? "bg-red-600 text-white" : "border border-red-200 bg-red-50 text-red-700 hover:bg-red-100"}`}>⚠️ Mark Risky</button>
                          )
                        ) : pickedStatus.isWeak ? (
                          <button type="button" onClick={() => setPickedMode(pickedMode === "clear" ? "normal" : "clear")} className={`inline-flex items-center gap-1 rounded px-2 py-0.5 text-[10px] font-bold transition ${pickedMode === "clear" ? "bg-blue-600 text-white" : "border border-blue-300 bg-blue-50 text-blue-800 hover:bg-blue-100"}`}><UserCheck size={10} /> Clear Weak Mark</button>
                        ) : (
                          <button type="button" onClick={() => setPickedMode(pickedMode === "weak" ? "normal" : "weak")} className={`inline-flex items-center gap-1 rounded px-2 py-0.5 text-[10px] font-bold transition ${pickedMode === "weak" ? "bg-teal-600 text-white" : "border border-teal-200 bg-teal-50 text-teal-800 hover:bg-teal-100"}`}>🌱 Weak to Nourish</button>
                        )}
                      </div>
                    </div>

                    {showMentionList && (
                      <div className="absolute right-0 top-7 z-40 w-56 max-h-48 overflow-y-auto rounded-xl border border-slate-200 bg-white p-1 shadow-xl space-y-0.5">
                        {allUsers.map((u: any) => {
                          const uName = u.full_name || u.email?.split("@")[0] || "Officer";
                          return (
                            <button
                              key={u.id}
                              type="button"
                              onClick={() => handlePickedTagUser(uName)}
                              className="w-full flex items-center gap-2 px-2 py-1.5 text-left text-xs hover:bg-blue-50 rounded-lg transition"
                            >
                              <UserAvatar name={uName} avatarUrl={u.avatar_url} size="w-5 h-5" textSize="text-[9px]" />
                              <span className="font-semibold text-slate-700 truncate">{uName}</span>
                            </button>
                          );
                        })}
                      </div>
                    )}

                    <textarea ref={pickedTextareaRef} value={noteText} onChange={e => setNoteText(e.target.value)} placeholder={pickedMode === "risky" ? "Comment why risky..." : pickedMode === "weak" ? "Comment why weak..." : pickedMode === "clear" ? "Comment why clearing mark..." : "Write tracking message or @tag someone..."} className={`min-h-[115px] flex-1 resize-none rounded-lg border px-3 py-2 text-sm outline-none transition ${pickedMode === "risky" ? "border-red-400 bg-red-50/40 focus:border-red-600" : pickedMode === "weak" ? "border-teal-400 bg-teal-50/40 focus:border-teal-600" : pickedMode === "clear" ? "border-blue-500 bg-blue-50/60 focus:border-blue-700" : "border-slate-300 bg-white focus:border-blue-500"}`} />
                    <div className="mt-2 flex items-center justify-between gap-2">
                      <span className="text-[11px] text-slate-500">By <b className="text-slate-700">{currentUserName}</b></span>
                      <button onClick={handleSavePicked} disabled={saving || !noteText.trim()} className={`flex items-center gap-1.5 rounded-lg px-4 py-1.5 text-xs font-bold text-white disabled:opacity-50 ${pickedMode === "risky" ? "bg-red-600 hover:bg-red-700" : pickedMode === "weak" ? "bg-teal-600 hover:bg-teal-700" : "bg-blue-600 hover:bg-blue-700"}`}><Send size={12} /> {saving ? "Saving..." : pickedMode === "risky" ? "Save Risky & Notify" : pickedMode === "weak" ? "Save Weak & Notify" : pickedMode === "clear" ? "Clear Mark" : "Save Tracking"}</button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </section>

      <div className={`grid transition-all duration-300 ease-in-out ${trackOpen ? "grid-rows-[1fr] opacity-100" : "grid-rows-[0fr] opacity-0 pointer-events-none"}`}>
        <div className="overflow-hidden">
          <div className="flex items-center justify-between rounded-xl border border-dashed border-slate-300 bg-slate-100/80 px-4 py-2.5 text-xs text-slate-600">
            <span>📋 Tracking list ({logGroups.length} agents) is minimized..</span>
            <button type="button" onClick={() => setTrackOpen(false)} className="font-bold text-blue-700 hover:underline shrink-0 ml-2">Show Messages ▾</button>
          </div>
        </div>
      </div>

      {/* 2. SMOOTH ANIMATED FILTERS & MESSAGES SECTION */}
      <div className={`grid transition-all duration-300 ease-in-out ${!trackOpen ? "grid-rows-[1fr] opacity-100 translate-y-0" : "grid-rows-[0fr] opacity-0 -translate-y-2 pointer-events-none"}`}>
        <div className="overflow-hidden">
          <section className="space-y-3 pt-0.5">
            <div className="rounded-xl border border-slate-200 bg-white p-2.5 sm:p-3 shadow-2xs space-y-2.5">
              <div className="flex flex-nowrap gap-1.5 rounded-xl bg-slate-100 p-1 w-full overflow-hidden">
                <button type="button" onClick={() => setStatusTab("all")} className={`flex items-center justify-center gap-1.5 rounded-lg py-1.5 px-2 text-xs font-bold transition-all duration-300 ${statusTab === "all" ? "bg-white text-slate-900 shadow-2xs shrink-0" : "text-slate-600 hover:text-slate-900 flex-1 min-w-0"}`}><span className={`whitespace-nowrap ${statusTab === "all" ? "" : "truncate"}`}>All Tracked</span></button>
                <button type="button" onClick={() => setStatusTab("risky")} className={`flex items-center justify-center gap-1.5 rounded-lg py-1.5 px-2 text-xs font-bold transition-all duration-300 ${statusTab === "risky" ? "bg-red-600 text-white shadow-2xs shrink-0" : "text-red-700 hover:bg-red-50 flex-1 min-w-0"}`}><AlertTriangle size={12} className="shrink-0" /><span className={`whitespace-nowrap ${statusTab === "risky" ? "" : "truncate"}`}>Risky Agent</span><span className={`rounded-full px-1.5 py-px text-[10px] shrink-0 ${statusTab === "risky" ? "bg-red-700 text-white" : "bg-red-100 text-red-700"}`}>{riskyTotalCount}</span></button>
                <button type="button" onClick={() => setStatusTab("weak")} className={`flex items-center justify-center gap-1.5 rounded-lg py-1.5 px-2 text-xs font-bold transition-all duration-300 ${statusTab === "weak" ? "bg-teal-600 text-white shadow-2xs shrink-0" : "text-teal-800 hover:bg-teal-50 flex-1 min-w-0"}`}><Sprout size={12} className="shrink-0" /><span className={`whitespace-nowrap ${statusTab === "weak" ? "" : "truncate"}`}>Weak to Nourish</span><span className={`rounded-full px-1.5 py-px text-[10px] shrink-0 ${statusTab === "weak" ? "bg-teal-700 text-white" : "bg-teal-100 text-teal-800"}`}>{weakTotalCount}</span></button>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex flex-wrap items-center gap-1.5">
                  <div className="inline-flex rounded-lg bg-slate-100 p-0.5 text-[11px] font-semibold">
                    <button type="button" onClick={() => { setDateMode("all"); setDateFrom(""); setDateTo(""); }} className={`rounded-md px-2.5 py-1 transition ${dateMode === "all" ? "bg-blue-600 text-white shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}>All Messages</button>
                    <button type="button" onClick={() => { setDateMode("single"); setDateTo(""); }} className={`rounded-md px-2.5 py-1 transition ${dateMode === "single" ? "bg-white text-slate-900 shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}>Single Date</button>
                    <button type="button" onClick={() => setDateMode("range")} className={`rounded-md px-2.5 py-1 transition ${dateMode === "range" ? "bg-white text-slate-900 shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}>Multi Date</button>
                  </div>
                  {dateMode === "single" && <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} className="rounded-lg border border-slate-300 bg-white px-2.5 py-1 text-xs outline-none focus:border-blue-500" />}
                  {dateMode === "range" && <div className="flex items-center gap-1"><input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} className="rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs outline-none focus:border-blue-500" /><span className="text-xs text-slate-400">to</span><input type="date" value={dateTo} min={dateFrom || undefined} onChange={e => setDateTo(e.target.value)} className="rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs outline-none focus:border-blue-500" /></div>}
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-[11px] font-semibold text-slate-500">{logGroups.length} agent(s)</span>
                  {hasLogFilter && <button type="button" onClick={resetLogFilters} className="inline-flex items-center gap-1 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 px-2.5 py-1 text-[11px] font-semibold text-slate-700"><RotateCcw size={12} /> Reset</button>}
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
                <div className="relative col-span-2 sm:col-span-1"><Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" /><input value={logSearch} onChange={e => setLogSearch(e.target.value)} placeholder="Agent, mobile, message..." className={`${selectCls} pl-7`} /></div>
                <div className="relative"><User size={12} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" /><select className={`${selectCls} pl-7`} value={logUser} onChange={e => setLogUser(e.target.value)}><option value="">All Users</option>{trackedUsers.map(u => <option key={u} value={u}>{u}</option>)}</select></div>
                <select className={selectCls} value={logDistrict} onChange={e => { setLogDistrict(e.target.value); setLogUpazila(""); }}><option value="">All Districts</option>{selDistricts.map(x => <option key={x} value={x}>{x}</option>)}</select>
                <select className={selectCls} value={logUpazila} onChange={e => setLogUpazila(e.target.value)}><option value="">All Upazilas</option>{logUpazilas.map(x => <option key={x} value={x}>{x}</option>)}</select>
              </div>
            </div>

            {loading ? <div className="rounded-xl border border-slate-200 bg-white p-8 text-center text-xs text-slate-400">Loading tracking messages...</div> : logGroups.length === 0 ? <div className="rounded-xl border border-slate-200 bg-white p-8 text-center text-xs text-slate-400">{statusTab === "risky" ? "No Risky Agents found." : statusTab === "weak" ? "No 'Weak to Nourish' agents found." : hasLogFilter ? "No tracking messages match your filters." : "No tracking messages yet. Tap 'Track an Agent' above to add one."}</div> : (
              <div className="space-y-2.5">
                {logGroups.slice(0, visibleCount).map(({ agent, notes }: any) => (
                  <AgentTrackingGroupCard
                    key={agent.id}
                    agent={agent}
                    notes={notes}
                    currentUserId={currentUserId}
                    currentUserName={currentUserName}
                    isFullAccess={isFullAccess}
                    getUserAvatar={getUserAvatar}
                    onTagFromNote={handleTagFromGroupCard}
                    onRetrack={() => setModal({ agent, tab: "track" })}
                    onDetails={() => setModal({ agent, tab: "details" })}
                    onViewMap={() => openMap(agent)}
                    onEditNote={updateNote}
                  />
                ))}
                {logGroups.length > visibleCount && <div className="pt-1 text-center"><button type="button" onClick={() => setVisibleCount((c) => c + 20)} className="rounded-xl border border-slate-300 bg-white hover:bg-slate-50 px-5 py-2 text-xs font-bold text-slate-700 shadow-2xs">Load more agents ({logGroups.length - visibleCount} remaining)</button></div>}
              </div>
            )}
          </section>
        </div>
      </div>

      {showPicker && <MapPickerModal agents={agents} initDistrict={selDistrict} initUpazila={selUpazila} onPick={pickAgent} onClose={() => setShowPicker(false)} />}
      {modal && (
        <AgentModal
          key={`${modal.agent.id}-${modal.tab}-${modal.tagUser || ""}`}
          agent={modal.agent}
          notes={notesByAgent.get(String(modal.agent.id)) || []}
          initialTab={modal.tab}
          initialTagUser={modal.tagUser}
          userId={currentUserId}
          userName={currentUserName}
          isFullAccess={isFullAccess}
          allUsers={allUsers}
          getUserAvatar={getUserAvatar}
          onSave={saveNote}
          onEditNote={updateNote}
          onViewMap={openMap}
          onClose={() => setModal(null)}
        />
      )}
      {mapAgent && <AgentMapModal agent={mapAgent} onClose={() => setMapAgent(null)} onDetails={() => { setModal({ agent: mapAgent, tab: "details" }); setMapAgent(null); }} />}
      {toast && <div className="fixed bottom-4 right-4 z-[3000] flex items-center gap-2 rounded-lg bg-slate-900 px-3.5 py-2 text-xs font-semibold text-white shadow-lg"><CheckCircle2 size={14} className="text-emerald-400" /> {toast}</div>}
    </div>
  );
}