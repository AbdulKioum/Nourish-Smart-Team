import React, { useEffect, useState, useRef } from "react";
import { supabase } from "../../lib/supabase";
import LocationPicker from "../../components/maps/LocationPicker";
import {
  getFarmers,
  createFarmer,
  updateFarmer,
  deleteFarmer,
} from "../../services/farmerService";
import { useAuth } from "../../context/AuthContext";
import {
  canEditRecord,
  canAddFarmer,
} from "../../utils/accessControl";
import * as XLSX from "xlsx";
import {
  MapContainer,
  Marker,
  Popup,
  Tooltip,
  useMap,
  Polyline,
} from "react-leaflet";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import MapLayers from "../../components/maps/MapLayers";
import {
  MapPin,
  Plus,
  Search,
  RotateCcw,
  FileSpreadsheet,
  BarChart3,
  Map as MapIcon,
  Send,
  X,
  CheckCircle2,
  Navigation,
} from "lucide-react";
import ExcelJS from "exceljs";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Capacitor } from "@capacitor/core";
import { Geolocation } from "@capacitor/geolocation";
import SearchableSelect from "../../components/common/SearchableSelect";
import { UserAvatar, renderNoteWithMentions } from "../../utils/avatarHelper";

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

function FitBounds({
  locations,
  skip,
}: {
  locations: { lat: number; lng: number }[];
  skip?: boolean;
}) {
  const map = useMap();

  useEffect(() => {
    if (skip) return;
    if (!locations || locations.length === 0) return;

    const bounds = L.latLngBounds(locations.map((loc) => [loc.lat, loc.lng]));
    map.fitBounds(bounds, { padding: [50, 50], maxZoom: 15 });
  }, [locations, map, skip]);

  return null;
}

function MapFocusController({
  farmerPos,
  myPos,
  myLocationMode,
}: {
  farmerPos: [number, number] | null;
  myPos: [number, number] | null;
  myLocationMode: boolean;
}) {
  const map = useMap();

  useEffect(() => {
    if (myPos && farmerPos) {
      const bounds = L.latLngBounds([myPos, farmerPos]);
      map.fitBounds(bounds, { padding: [65, 65], maxZoom: 16, animate: true });
    } else if (farmerPos) {
      map.flyTo(farmerPos, 17, { animate: true, duration: 1.0 });
    } else if (myLocationMode && myPos) {
      map.flyTo(myPos, 16, { animate: true, duration: 1.0 });
    }
  }, [farmerPos, myPos, myLocationMode, map]);

  return null;
}

delete (L.Icon.Default.prototype as any)._getIconUrl;

L.Icon.Default.mergeOptions({
  iconRetinaUrl:
    "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png",
  iconUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png",
  shadowUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png",
});

const EMPTY_FORM = {
  company_id: "",
  agent_id: "",
  farmer_name: "",
  mobile: "",
  village: "",
  union_name: "",
  upazila: "",
  district: "",
  farm_category: "",
  broiler: "",
  layer: "",
  sonali: "",
  duck: "",
  quail: "",
  deshi_chicken: "",
  dairy_cow: "",
  bull: "",
  calf_heifer: "",
  pangus: "",
  tilapia: "",
  koi: "",
  shing: "",
  pabda: "",
  gulsha: "",
  magur: "",
  carp: "",
  golda: "",
  bagda: "",
  tengra: "",
  feed_type: "",
  latitude: "",
  longitude: "",
  status: "Active",
  remarks: "",
};

const STOCK_FIELDS = [
  "broiler","layer","sonali","duck","quail","deshi_chicken",
  "dairy_cow","bull","calf_heifer",
  "pangus","tilapia","koi","shing","pabda","gulsha","magur","carp","golda","bagda","tengra",
];

const EMPTY_SHED_FORM: any = {
  id: undefined,
  shed_name: "",
  farm_category: "",
  broiler: "", layer: "", sonali: "", duck: "", quail: "", deshi_chicken: "",
  dairy_cow: "", bull: "", calf_heifer: "",
  pangus: "", tilapia: "", koi: "", shing: "", pabda: "", gulsha: "",
  magur: "", carp: "", golda: "", bagda: "", tengra: "",
  feed_type: "",
  latitude: "", longitude: "",
  status: "Active",
};

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

function getFarmerBadgeInfo(farmerOrCategory: any) {
  // Backward compat: accepts a farmer/shed object OR a category string.
  const farmer =
    typeof farmerOrCategory === "string"
      ? { farm_category: farmerOrCategory }
      : farmerOrCategory || {};
  if (farmer.farm_category === "Poultry") {
    if (Number(farmer.layer || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🥚", label: "Layer", bg: "#d97706", border: "#b45309" };
    }
    if (Number(farmer.sonali || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🐓", label: "Sonali", bg: "#ea580c", border: "#c2410c" };
    }
    if (Number(farmer.duck || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🦆", label: "Duck", bg: "#0d9488", border: "#0f766e" };
    }
    return { emoji: "🐔", label: "Broiler", bg: "#f59e0b", border: "#d97706" };
  }
  if (farmer.farm_category === "Cattle") {
    return { emoji: "🐄", label: "Cattle", bg: "#16a34a", border: "#15803d" };
  }
  if (farmer.farm_category === "Fish") {
    return { emoji: "🐟", label: "Fish", bg: "#0284c7", border: "#0369a1" };
  }
  return { emoji: "🏡", label: "Farm", bg: "#475569", border: "#334155" };
}

// Clean Marker without circular background; auto-scales horizontal vs square/round logos
function buildCustomFarmMarker(
  farmer: any,
  shed: any = null,
  highlight: any = false,
  distText?: string,
  shedIndex: number = 0,
  totalSheds: number = 1
) {
  // Backward compat: old calls passed (farmer, highlight, distText).
  if (typeof shed === "boolean") {
    distText = typeof highlight === "string" ? highlight : distText;
    highlight = shed;
    shed = null;
  }
  if (typeof highlight === "string") {
    distText = highlight;
    highlight = false;
  }
  const badgeSource = shed || farmer;
  const { emoji, border } = getFarmerBadgeInfo(badgeSource);
  const logoUrl = farmer.companies?.logo_url;
  const labelText = shed
    ? getShedDisplayName(farmer, shed, shedIndex, totalSheds)
    : farmer.farmer_name || "Farmer";
  const safeName = String(labelText).replace(/</g, "&lt;").slice(0, 20);

  const ring = highlight
    ? "0 0 0 3px rgba(16,185,129,0.45)"
    : "0 1px 2px rgba(0,0,0,0.15)";

  const distBadge = distText
    ? `<div style="margin-top:2px;background:#0f172a;color:#ffffff;font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;white-space:nowrap;">${distText}</div>`
    : "";

  const html = `
    <div style="display:flex;flex-direction:column;align-items:center;transform:translateY(-14px);">
      <div style="position:relative;display:flex;flex-direction:column;align-items:center;justify-content:center;">
        <span style="font-size:${highlight ? 25 : 20}px;line-height:1;filter:drop-shadow(0 2px 3px rgba(0,0,0,0.35));">${emoji}</span>
        ${
          logoUrl
            ? `<img src="${logoUrl}" onload="if(this.naturalWidth > this.naturalHeight * 1.3){this.style.width='44px';this.style.height='auto';this.style.maxHeight='20px';}else{this.style.width='22px';this.style.height='22px';}" style="margin-top:1px;width:24px;height:22px;object-fit:contain;background:rgba(255,255,255,0.92);padding:1px 2px;border-radius:3px;box-shadow:0 1px 2px rgba(0,0,0,0.2);" />`
            : ""
        }
      </div>
      <div style="margin-top:2px;background:rgba(252, 244, 236, 0.95);border:1.5px solid ${border};color:#314158;font-size:10px;font-weight:700;padding:0px 5px;border-radius:4px;white-space:nowrap;box-shadow:${ring};">
        ${safeName}
      </div>
      ${distBadge}
    </div>
  `;

  return L.divIcon({
    className: "custom-farm-div-icon",
    html,
    iconSize: [52, 54],
    iconAnchor: [26, 44],
    popupAnchor: [0, -38],
  });
}

function getStockBreakdownList(farmer: any): { label: string; qty: number }[] {
  const list: { label: string; qty: number }[] = [];
  const allKeys = [...POULTRY_FIELDS, ...CATTLE_FIELDS, ...FISH_FIELDS];
  allKeys.forEach((k) => {
    const val = Number(farmer?.[k] || 0);
    if (val > 0) {
      const label = k
        .replace("_", " ")
        .replace(/\b\w/g, (c) => c.toUpperCase());
      list.push({ label, qty: val });
    }
  });
  return list;
}

// ---------- Multi-shed helpers ----------
function getShedsForFarmer(farmer: any): any[] {
  const list = farmer?.farmer_sheds || farmer?.sheds || [];
  if (Array.isArray(list) && list.length > 0) return list;
  // Legacy fallback: synthesise ONE shed from the farmer's own fields.
  // shed_name stays EMPTY — a single shed has no name, so the shared
  // short-name logic returns "" and the full name is just the farmer name.
  const shed: any = {
    id: `legacy-${farmer?.id}`,
    farmer_id: farmer?.id,
    shed_name: "",
    farm_category: farmer?.farm_category || "",
    latitude: farmer?.latitude ?? null,
    longitude: farmer?.longitude ?? null,
    feed_type: farmer?.feed_type || "",
    status: farmer?.status || "Active",
  };
  STOCK_FIELDS.forEach((k) => { shed[k] = Number(farmer?.[k] || 0); });
  return [shed];
}

// Shed naming — single source of truth, identical in Farmers/Visits/Doc:
// - Stored shed_name is the SHORT label only ("Shed 1", "Shed 2", "Bridge par").
// - A farmer with a single shed (no shed explicitly added) has NO shed name:
//   the Shed column is blank and the full name is just the farmer name.
// - getShedShortName() normalises any legacy stored value: empty, or exactly
//   the farmer name, or an auto "Shed N" label, is blank for a single shed
//   and "Shed N" once a 2nd shed exists; a value starting with the farmer
//   name (full name or first name, e.g. "Raju Ahmed - Shed 1" or "Raju Shed 1")
//   has that prefix stripped; "Shed-1" normalises to "Shed 1"; a bare custom
//   label ("Bridge par") is kept as-is.
// - getShedDisplayName() is just the farmer name for a single unnamed shed,
//   otherwise "Farmer Name - Short Name". The ONLY place showing the short
//   name alone is the farmer table Shed column (and the form Shed Name input).
function getShedShortName(
  farmer: any,
  shed: any,
  index: number = 0,
  totalSheds: number = 1
): string {
  const farmerName = String(farmer?.farmer_name || "").trim() || "Farmer";
  const raw = String(shed?.shed_name || "").trim();
  const lowerRaw = raw.toLowerCase();
  const lowerFarmer = farmerName.toLowerCase();
  const firstName = lowerFarmer.split(/\s+/)[0] || "";
  let rest = raw;
  const stripPrefix = (prefix: string): boolean => {
    if (prefix && lowerRaw.startsWith(prefix)) {
      const next = raw.charAt(prefix.length);
      if (next === " " || next === "-") {
        rest = raw.slice(prefix.length).replace(/^[\s-]+/, "").trim();
        return true;
      }
    }
    return false;
  };
  if (rest && rest.toLowerCase() !== lowerFarmer) {
    if (!stripPrefix(lowerFarmer) && firstName) stripPrefix(firstName);
  } else {
    rest = "";
  }
  // Auto-generated labels ("Shed 1", "Shed-1") carry no real name.
  const autoMatch = rest.match(/^shed\s*-?\s*(\d+)$/i);
  if ((totalSheds ?? 1) <= 1) {
    if (!rest || autoMatch) return "";
    return rest;
  }
  if (!rest) return `Shed ${index + 1}`;
  if (autoMatch) return `Shed ${autoMatch[1]}`;
  return rest;
}

function getShedDisplayName(
  farmer: any,
  shed: any,
  index: number = 0,
  totalSheds: number = 1
): string {
  const farmerName = String(farmer?.farmer_name || "").trim() || "Farmer";
  const short = getShedShortName(farmer, shed, index, totalSheds);
  return short ? `${farmerName} - ${short}` : farmerName;
}

// Days since a date (identical helper in VisitsPage).
function getDaysSince(dateStr?: string | null): number | null {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return null;
  return Math.floor((new Date().getTime() - d.getTime()) / 86400000);
}

// Per-shed status — the same rule as the farmer-level computedStatus in
// VisitsPage, but keyed by shed (identical helper in VisitsPage): a shed
// with a Running/Activated visit is Running; otherwise the status comes
// from that shed's own last sale (<=60 days => Active, older =>
// Inactive); a shed with no visit history falls back to its stored
// status (default Active). A visit without a shed_id counts only when
// the farmer has a single shed.
function getShedVisitStatus(
  farmer: any,
  shed: any,
  farmerVisits: any[]
): {
  status: "Running" | "Active" | "Inactive";
  statusDate: string;
  daysSinceSale: number | null;
  closedType: string;
} {
  const sheds = getShedsForFarmer(farmer);
  const shedId =
    shed?.id != null && !String(shed.id).startsWith("legacy-")
      ? String(shed.id)
      : null;
  const shedVisits = (farmerVisits || []).filter((v: any) => {
    const sid = v?.shed_id;
    if (sid !== null && sid !== undefined && String(sid) !== "") {
      return shedId !== null && String(sid) === shedId;
    }
    return sheds.length === 1;
  });
  const running = shedVisits.filter(
    (v: any) => v.status === "Activated" || v.status === "Running"
  );
  const closed = shedVisits
    .filter((v: any) => v.status === "Sold" || v.status === "Culled")
    .sort((a: any, b: any) =>
      String(b.closed_date || b.created_at || "").localeCompare(
        String(a.closed_date || a.created_at || "")
      )
    );
  const lastSale =
    closed[0]?.closed_date ||
    (closed[0]?.created_at ? String(closed[0].created_at).slice(0, 10) : null);
  const days = getDaysSince(lastSale);
  let status: "Running" | "Active" | "Inactive";
  if (running.length > 0) status = "Running";
  else if (days !== null && days <= 60) status = "Active";
  else if (days !== null) status = "Inactive";
  else
    status =
      shed?.status === "Running" ||
      shed?.status === "Active" ||
      shed?.status === "Inactive"
        ? shed.status
        : "Active";
  return {
    status,
    statusDate: lastSale || "",
    daysSinceSale: days,
    closedType: closed[0]?.status || "",
  };
}

function getFarmerCategories(farmer: any): string[] {
  const seen: string[] = [];
  getShedsForFarmer(farmer).forEach((s: any) => {
    if (s?.farm_category && !seen.includes(s.farm_category)) seen.push(s.farm_category);
  });
  return seen;
}

function getShedStockBreakdown(shed: any): { label: string; qty: number }[] {
  const list: { label: string; qty: number }[] = [];
  STOCK_FIELDS.forEach((k) => {
    const val = Number(shed?.[k] || 0);
    if (val > 0) {
      const label = k.replace("_", " ").replace(/\b\w/g, (c) => c.toUpperCase());
      list.push({ label, qty: val });
    }
  });
  return list;
}

function getFarmerShedsCount(farmer: any): number {
  return getShedsForFarmer(farmer).length;
}

function getTotalSheds(list: any[]): number {
  return (list || []).reduce((sum, f) => sum + getFarmerShedsCount(f), 0);
}

// Multi-company agents (agent_companies table): an agent matches a company
// if ANY linked company equals it. Falls back to the single company_id.
function getAgentCompanyIds(agent: any): string[] {
  const linked = agent?.linked_company_ids || agent?.agent_companies || [];
  if (Array.isArray(linked) && linked.length > 0) {
    return linked
      .map((c: any) =>
        String(typeof c === "object" ? c.company_id ?? c.id : c)
      )
      .filter(Boolean);
  }
  return agent?.company_id != null && agent.company_id !== ""
    ? [String(agent.company_id)]
    : [];
}

function isAgentNonExclusive(agent: any): boolean {
  return getAgentCompanyIds(agent).length > 1;
}



function shedFormToDb(shed: any, farmerId: number, farmerName: string, index: number, totalSheds: number = 1): any {
  const db: any = {
    farmer_id: farmerId,
    shed_name: getShedShortName({ farmer_name: farmerName }, shed, index, totalSheds),
    farm_category: shed?.farm_category || "",
    latitude: shed?.latitude === "" || shed?.latitude == null ? null : Number(shed.latitude),
    longitude: shed?.longitude === "" || shed?.longitude == null ? null : Number(shed.longitude),
    feed_type: shed?.feed_type || "",
    status: shed?.status || "Active",
  };
  STOCK_FIELDS.forEach((k) => { db[k] = Number(shed?.[k] || 0); });
  return db;
}

export default function FarmersPage() {
  const { role } = useAuth();
  const mapSectionRef = useRef<HTMLDivElement | null>(null);

  if (role === "viewer") {
    return <div className="p-6 text-center text-slate-500">Access Denied</div>;
  }

  const isFullAccess = role === "super_admin" || role === "manager";
  // Permissions (src/utils/accessControl.ts): manager / super admin can do
  // everything; officers edit/delete only their own rows (old rows with
  // created_by null are editable by area officers — the save stamps the
  // creator). Farmer add is NOT category-gated (rule 5). View / Map
  // stay open to everyone.
  
  const canAddFarmerHere = canAddFarmer(role);
  const canEditFarmer = (farmer: any) =>
    canEditRecord(role, farmer?.created_by, currentUserId);

  const [currentUserId, setCurrentUserId] = useState("");
  const [currentUserName, setCurrentUserName] = useState("User");
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);
  const [_assignedCategories, setAssignedCategories] = useState<string[]>([]);

  const [farmers, setFarmers] = useState<any[]>([]);
  // vet_visits rows (for per-shed status: a shed with a Running visit
  // is Running; otherwise its status comes from its own last sale).
  const [visitRows, setVisitRows] = useState<any[]>([]);
  const [companies, setCompanies] = useState<any[]>([]);
  const [agents, setAgents] = useState<any[]>([]);

  const [search, setSearch] = useState("");
  const [editingId, setEditingId] = useState<number | null>(null);
  const [selectedFarmer, setSelectedFarmer] = useState<any>(null);
  const [showMap, setShowMap] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [showFarmerMap, setShowFarmerMap] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showSummary, setShowSummary] = useState(false);
  const [form, setForm] = useState({ ...EMPTY_FORM });
  const [formSheds, setFormSheds] = useState<any[]>([{ ...EMPTY_SHED_FORM }]);
  const [pickingShedIndex, setPickingShedIndex] = useState<number | null>(null);
  const [_selectedShed, setSelectedShed] = useState<any>(null);

  const [pickerCenterOverride, setPickerCenterOverride] = useState<
    [number, number] | null
  >(null);
  const [locatingInPicker, setLocatingInPicker] = useState(false);

  const [activeTrackingFarmerId, setActiveTrackingFarmerId] = useState<
    number | null
  >(null);
  const [farmerNotes, setFarmerNotes] = useState<any[]>([]);
  const [newNoteText, setNewNoteText] = useState("");
  const [loadingNotes, setLoadingNotes] = useState(false);
  const [savingNote, setSavingNote] = useState(false);

  const [companyFilter, setCompanyFilter] = useState("");
  const [agentFilter, setAgentFilter] = useState("");
  const [districtFilter, setDistrictFilter] = useState("");
  const [upazilaFilter, setUpazilaFilter] = useState("");
  const [unionFilter, setUnionFilter] = useState("");
  const [farmCategoryFilter, setFarmCategoryFilter] = useState("");

  const [districts, setDistricts] = useState<any[]>([]);
  const [upazilas, setUpazilas] = useState<any[]>([]);
  const [unions, setUnions] = useState<any[]>([]);
  const [filterUpazilas, setFilterUpazilas] = useState<any[]>([]);
  const [filterUnions, setFilterUnions] = useState<any[]>([]);

  const [myLocationMode, setMyLocationMode] = useState(false);
  const [selectedFarmerPosition, setSelectedFarmerPosition] = useState<
    [number, number] | null
  >(null);
  const [myPosition, setMyPosition] = useState<[number, number] | null>(null);
  const farmerMarkerRefs = useRef<Record<string, L.Marker | null>>({});
  const myLocationMarkerRef = useRef<L.Marker | null>(null);

  const [allUsers, setAllUsers] = useState<any[]>([]);

  useEffect(() => {
    supabase
      .from("users")
      .select("id, full_name, avatar_url")
      .then(({ data }) => {
        if (data) setAllUsers(data);
      });
  }, []);

  function getUserAvatar(userId?: string, userName?: string): string | null {
    const found = allUsers.find(
      (u) =>
        (userId && String(u.id) === String(userId)) ||
        (userName &&
          String(u.full_name || "").trim().toLowerCase() ===
            String(userName).trim().toLowerCase())
    );
    return found?.avatar_url || null;
  }

  async function loadData() {
    const {
      data: { user },
    } = await supabase.auth.getUser();

    let uDistricts: string[] = [];
    let uUpazilas: string[] = [];
    let uName = user?.email?.split("@")[0] || "Officer";

    if (user) {
      setCurrentUserId(user.id);
      // assigned_categories comes from the permissions migration;
      // fall back gracefully if the column is not there yet.
      let profile: any = null;
      {
        const { data, error } = await supabase
          .from("users")
          .select("full_name, email, role, districts, upazilas, assigned_categories")
          .eq("id", user.id)
          .single();
        if (!error && data) {
          profile = data;
        } else {
          const retry = await supabase
            .from("users")
            .select("full_name, email, role, districts, upazilas")
            .eq("id", user.id)
            .single();
          profile = retry.data || null;
        }
      }

      if (profile) {
        uName = profile.full_name || profile.email || uName;
        uDistricts = parseList(profile.districts);
        uUpazilas = parseList(profile.upazilas);
        setAssignedCategories(parseList(profile.assigned_categories));
      }
    }

    setCurrentUserName(uName);
    setAssignedUpazilas(uUpazilas);

    const [
      farmerData,
      { data: companyData },
      { data: agentData },
      { data: districtData },
    ] = await Promise.all([
      getFarmers(),
      supabase.from("companies").select("*").order("company_name"),
      supabase.from("agents").select("*").order("agent_name"),
      supabase.from("districts").select("*").order("name"),
    ]);

    const allFarmers = farmerData || [];
    // Attach sheds from farmer_sheds when the table exists; otherwise
    // getShedsForFarmer() synthesises one legacy shed per farmer.
    try {
      const { data: shedData, error: shedError } = await supabase
        .from("farmer_sheds")
        .select("*");
      if (!shedError && shedData) {
        const shedsByFarmer: Record<string, any[]> = {};
        shedData.forEach((s: any) => {
          const k = String(s.farmer_id);
          if (!shedsByFarmer[k]) shedsByFarmer[k] = [];
          shedsByFarmer[k].push(s);
        });
        allFarmers.forEach((f: any) => {
          if (shedsByFarmer[String(f.id)]) f.farmer_sheds = shedsByFarmer[String(f.id)];
        });
      }
    } catch {
      // farmer_sheds table not created yet — legacy fallback applies.
    }
    // Ownership: attach created_by when the column exists (the farmer
    // service may not select it). A missing column must not break the
    // list — rows without a creator stay editable (rule 2, old data).
    try {
      const { data: creatorData, error: creatorErr } = await supabase
        .from("farmers")
        .select("id, created_by");
      if (!creatorErr && creatorData) {
        const creatorById: Record<string, any> = {};
        (creatorData as any[]).forEach((r: any) => {
          creatorById[String(r.id)] = r.created_by;
        });
        allFarmers.forEach((f: any) => {
          if (f.created_by == null && creatorById[String(f.id)] !== undefined)
            f.created_by = creatorById[String(f.id)];
        });
      }
    } catch {
      // created_by column not added yet — old rows stay editable.
    }

    const allowedFarmers = isFullAccess
      ? allFarmers
      : allFarmers.filter((f: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(f.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(f.district);
          return matchUpa && matchDist;
        });

    setFarmers(allowedFarmers);

    // Per-shed status: load vet visits so each shed row can show its
    // own Running/Active/Inactive status (same rule as VisitsPage).
    // Best effort — if the query fails, shed rows fall back to the
    // stored shed.status.
    try {
      const { data: visitData, error: visitErr } = await supabase
        .from("vet_visits")
        .select("id, farmer_id, shed_id, status, closed_date, created_at");
      if (!visitErr && visitData) setVisitRows(visitData as any[]);
    } catch {
      // vet_visits not readable — stored shed status applies.
    }
    setCompanies(companyData || []);

    const allAgents = agentData || [];
    // Multi-company agents: attach linked company ids (agent_companies).
    // If the table is missing, the single company_id fallback applies.
    try {
      const { data: linkData, error: linkError } = await supabase
        .from("agent_companies")
        .select("agent_id, company_id");
      if (!linkError && linkData) {
        const byAgent: Record<string, string[]> = {};
        (linkData as any[]).forEach((l: any) => {
          const k = String(l.agent_id);
          if (!byAgent[k]) byAgent[k] = [];
          if (!byAgent[k].includes(String(l.company_id)))
            byAgent[k].push(String(l.company_id));
        });
        (allAgents as any[]).forEach((a: any) => {
          if (byAgent[String(a.id)]) a.linked_company_ids = byAgent[String(a.id)];
        });
      }
    } catch {
      // agent_companies table not created yet — single company_id fallback applies.
    }

    const allowedAgents = isFullAccess
      ? allAgents
      : allAgents.filter((a: any) => {
          const aDist = a.district || a.district_name || "";
          const aUpa = a.upazila || a.upazila_name || "";
          const matchUpa =
            uUpazilas.length === 0 || !aUpa || uUpazilas.includes(aUpa);
          const matchDist =
            uDistricts.length === 0 || !aDist || uDistricts.includes(aDist);
          const hasAllowedFarmer = allowedFarmers.some(
            (f: any) => String(f.agent_id) === String(a.id)
          );
          if (aUpa || aDist) return matchUpa && matchDist;
          return hasAllowedFarmer;
        });

    setAgents(allowedAgents);

    const allDistricts = districtData || [];
    const allowedDists = (
      isFullAccess
        ? allDistricts
        : allDistricts.filter(
            (d: any) => uDistricts.length === 0 || uDistricts.includes(d.name)
          )
    ).sort((a: any, b: any) => (a.name || "").localeCompare(b.name || ""));

    setDistricts(allowedDists);

    if (!districtFilter) {
      if (allowedDists.length === 1) {
        setDistrictFilter(allowedDists[0].name);
      } else if (allowedDists.length > 0 && !isFullAccess) {
        setDistrictFilter(allowedDists[0].name);
      } else if (
        allowedDists.some((d: any) => d.name === "Mymensingh") &&
        isFullAccess
      ) {
        setDistrictFilter("Mymensingh");
      }
    }
  }

  useEffect(() => {
    loadData();
  }, [role]);

  function clearFilters() {
    setCompanyFilter("");
    setAgentFilter("");
    setDistrictFilter(districts.length === 1 ? districts[0].name : "");
    setUpazilaFilter(
      !isFullAccess && assignedUpazilas.length === 1 ? assignedUpazilas[0] : ""
    );
    setUnionFilter("");
    setFarmCategoryFilter("");
    setSearch("");
  }

  async function loadFarmerNotes(farmerId: number) {
    setActiveTrackingFarmerId(farmerId);
    setLoadingNotes(true);
    try {
      const { data, error } = await supabase
        .from("farmer_trackings")
        .select("*")
        .eq("farmer_id", farmerId)
        .order("created_at", { ascending: true });

      if (!error) {
        setFarmerNotes(data || []);
      } else {
        setFarmerNotes([]);
      }
    } catch {
      setFarmerNotes([]);
    } finally {
      setLoadingNotes(false);
    }
  }

  useEffect(() => {
    if (selectedFarmer?.id) {
      setNewNoteText("");
      loadFarmerNotes(selectedFarmer.id);
    }
  }, [selectedFarmer]);

  async function handleAddFarmerNote(
    e: React.FormEvent,
    targetFarmerId?: number
  ) {
    e.preventDefault();
    const fId = targetFarmerId || selectedFarmer?.id || activeTrackingFarmerId;
    if (!fId || !newNoteText.trim()) return;

    setSavingNote(true);
    try {
      const { error } = await supabase.from("farmer_trackings").insert([
        {
          farmer_id: fId,
          note: newNoteText.trim(),
          user_id: currentUserId || null,
          user_name: currentUserName,
        },
      ]);

      if (error) throw error;
      setNewNoteText("");
      await loadFarmerNotes(fId);
    } catch (err: any) {
      alert(err?.message || "Failed to save tracking note.");
    } finally {
      setSavingNote(false);
    }
  }

  // Map button: zoom to a specific shed when given, else the first shed with coords.
  function handleViewFarmerOnMap(farmer: any, shed?: any) {
    const targetShed =
      shed ||
      getShedsForFarmer(farmer).find(
        (s: any) => s.latitude && s.longitude
      );
    if (!targetShed || !targetShed.latitude || !targetShed.longitude) {
      alert("No map coordinates saved for this shed.");
      return;
    }
    const lat = Number(targetShed.latitude);
    const lng = Number(targetShed.longitude);
    setShowSummary(false);
    setShowFarmerMap(true);
    setMyLocationMode(false);
    setSelectedFarmerPosition([lat, lng]);
    setSelectedShed(targetShed);
    loadFarmerNotes(farmer.id);
    Object.values(farmerMarkerRefs.current).forEach((m) => m?.closePopup());
    setTimeout(() => {
      mapSectionRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "center",
      });
    }, 300);
  }

  const selectedDistrict = districts.find((d) => d.name === form.district);

  useEffect(() => {
    async function loadUpazilas() {
      if (!selectedDistrict) {
        setUpazilas([]);
        return;
      }
      const { data } = await supabase
        .from("upazilas")
        .select("*")
        .eq("district_id", selectedDistrict.id)
        .order("name");

      const list = data || [];
      const allowed = isFullAccess
        ? list
        : list.filter(
            (u: any) =>
              assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
          );
      setUpazilas(allowed);

      if (
        allowed.length > 0 &&
        (!form.upazila || !allowed.some((u: any) => u.name === form.upazila))
      ) {
        setForm((prev) => ({ ...prev, upazila: allowed[0].name }));
      }
    }
    loadUpazilas();
  }, [selectedDistrict, isFullAccess, assignedUpazilas]);

  const selectedUpazila = upazilas.find((u) => u.name === form.upazila);

  // Map picker centre: the farmer's agent first (farmers cluster around
  // their agent); falls back to upazila/district when the agent has no lat/lon.
  const formAgent = form.agent_id
    ? agents.find((a: any) => String(a.id) === String(form.agent_id))
    : null;
  const formAgentCenter: [number, number] | null =
    formAgent?.latitude && formAgent?.longitude
      ? [Number(formAgent.latitude), Number(formAgent.longitude)]
      : null;

  useEffect(() => {
    async function loadUnions() {
      if (!selectedUpazila) {
        setUnions([]);
        return;
      }
      const { data } = await supabase
        .from("unions")
        .select("*")
        .eq("upazila_id", selectedUpazila.id)
        .order("name");
      setUnions(data || []);
    }
    loadUnions();
  }, [selectedUpazila]);

  function getAgentLinkedCompanyNames(agent: any): string[] {
    return getAgentCompanyIds(agent)
      .map(
        (id: string) =>
          companies.find((c: any) => String(c.id) === id)?.company_name
      )
      .filter(Boolean);
  }

  const formAvailableAgents = agents.filter((a: any) => {
    if (
      form.company_id &&
      !getAgentCompanyIds(a).includes(String(form.company_id))
    ) {
      return false;
    }
    const aUpa = a.upazila || a.upazila_name || "";
    if (form.upazila && aUpa && aUpa !== form.upazila) {
      return false;
    }
    return true;
  });

  useEffect(() => {
    async function loadFilterUpazilas() {
      let query = supabase.from("upazilas").select("*").order("name");

      if (districtFilter) {
        const district = districts.find((d) => d.name === districtFilter);
        if (!district) {
          setFilterUpazilas([]);
          return;
        }
        query = query.eq("district_id", district.id);
      } else if (!isFullAccess && districts.length > 0) {
        const distIds = districts.map((d) => d.id);
        query = query.in("district_id", distIds);
      }

      const { data } = await query;
      const list = data || [];
      const allowed = isFullAccess
        ? list
        : list.filter(
            (u: any) =>
              assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
          );
      setFilterUpazilas(allowed);
    }
    loadFilterUpazilas();
  }, [districtFilter, districts, isFullAccess, assignedUpazilas]);

  useEffect(() => {
    async function loadFilterUnions() {
      if (upazilaFilter) {
        const upazila = filterUpazilas.find((u) => u.name === upazilaFilter);
        if (!upazila) {
          setFilterUnions([]);
          return;
        }
        const { data } = await supabase
          .from("unions")
          .select("*")
          .eq("upazila_id", upazila.id)
          .order("name");
        setFilterUnions(data || []);
        return;
      }

      const upazilaIds = filterUpazilas.map((u: any) => u.id);
      if (upazilaIds.length === 0) {
        setFilterUnions([]);
        return;
      }

      const { data } = await supabase
        .from("unions")
        .select("*")
        .in("upazila_id", upazilaIds)
        .order("name");
      setFilterUnions(data || []);
    }
    loadFilterUnions();
  }, [upazilaFilter, districtFilter, filterUpazilas]);

  const availableFilterAgents = agents.filter((a: any) => {
    if (
      companyFilter &&
      !getAgentCompanyIds(a).includes(String(companyFilter))
    ) {
      return false;
    }
    const aDist = a.district || a.district_name || "";
    const aUpa = a.upazila || a.upazila_name || "";

    if (districtFilter) {
      if (aDist && aDist !== districtFilter) return false;
      if (!aDist) {
        const hasFarmerInDist = farmers.some(
          (f) =>
            String(f.agent_id) === String(a.id) && f.district === districtFilter
        );
        if (!hasFarmerInDist) return false;
      }
    }

    if (upazilaFilter) {
      if (aUpa && aUpa !== upazilaFilter) return false;
      if (!aUpa) {
        const hasFarmerInUpa = farmers.some(
          (f) =>
            String(f.agent_id) === String(a.id) && f.upazila === upazilaFilter
        );
        if (!hasFarmerInUpa) return false;
      }
    }

    return true;
  });

  function updateShedField(index: number, field: string, value: any) {
    setFormSheds((prev) =>
      prev.map((s, i) => (i === index ? { ...s, [field]: value } : s))
    );
  }

  function addShed() {
    // Adding a 2nd shed: if the only existing shed has no name yet, ask for
    // its name first. Empty/cancel => it becomes "Shed 1" automatically.
    // The previous shed's name is saved to the DB right away (when it is an
    // existing shed row), not just in the form state.
    const prev = formSheds && formSheds.length > 0 ? formSheds[0] : null;
    const prevNeedsName =
      formSheds.length === 1 &&
      prev &&
      getShedShortName({ farmer_name: form.farmer_name }, prev, 0, 1) === "";
    let prevName = "";
    if (prevNeedsName) {
      const answer = window.prompt(
        "Enter name for the shed or will be shed 1:",
        ""
      );
      prevName =
        answer && answer.trim()
          ? getShedShortName({ farmer_name: form.farmer_name }, { shed_name: answer.trim() }, 0, 2)
          : "Shed 1";
      if (prev && prev.id != null && !String(prev.id).startsWith("legacy-")) {
        const prevId = Number(prev.id);
        supabase
          .from("farmer_sheds")
          .update({ shed_name: prevName })
          .eq("id", prevId)
          .then(({ error }: any) => {
            if (error) console.error("Could not save previous shed name:", error);
          });
      }
    }
    // Prefill the SHORT label only ("Shed 2"); the full name
    // ("Farmer Name - Shed 2") is composed at display time.
    setFormSheds((prevSheds) => {
      const newCount = prevSheds.length + 1;
      const updated = prevNeedsName
        ? prevSheds.map((s, i) => (i === 0 ? { ...s, shed_name: prevName } : s))
        : prevSheds;
      return [...updated, { ...EMPTY_SHED_FORM, shed_name: `Shed ${newCount}` }];
    });
  }

  function removeShed(index: number) {
    setFormSheds((prev) => (prev.length <= 1 ? prev : prev.filter((_, i) => i !== index)));
  }

  async function handleLocateMeForShed(index: number) {
    setLocatingInPicker(true);
    try {
      const { lat, lng } = await getDeviceCoordinates();
      setPickerCenterOverride([lat, lng]);
      setFormSheds((prev) =>
        prev.map((s, i) =>
          i === index ? { ...s, latitude: lat.toFixed(6), longitude: lng.toFixed(6) } : s
        )
      );
    } catch (err: any) {
      alert(err?.message || "Unable to retrieve your location. Please allow GPS access.");
    } finally {
      setLocatingInPicker(false);
    }
  }

  async function handleSave() {
    // Permission guard (buttons are hidden too — this is the backstop
    // for the save itself). Farmer add is NOT category-gated (rule 5).
    if (editingId) {
      const editingFarmer = farmers.find(
        (f: any) => Number(f.id) === Number(editingId)
      );
      if (!canEditFarmer(editingFarmer || {})) {
        alert("You can only edit farmers you added.");
        return;
      }
    } else if (!canAddFarmerHere) {
      alert("You do not have permission to add farmers.");
      return;
    }
    if (
      !form.company_id ||
      !form.agent_id ||
      !form.farmer_name.trim() ||
      !form.district ||
      !form.upazila
    ) {
      alert("Company, District, Upazila, Agent and Farmer Name are required");
      return;
    }
    const sheds = formSheds && formSheds.length > 0 ? formSheds : [{ ...EMPTY_SHED_FORM }];
    if (sheds.some((s: any) => !s.farm_category)) {
      alert("Every shed must have a Farm Category.");
      return;
    }
    // Shed names are stored as SHORT labels only — shedFormToDb strips any
    // farmer-name prefix. A single shed stays blank (no "Shed 1"); with 2+
    // sheds a blank label becomes "Shed N".
    const shedsWithNames = sheds;
    const firstShed = shedsWithNames[0];
    // Legacy farmer payload: first shed's category/location, stock SUMMED across sheds.
    const summedStock: Record<string, number> = {};
    STOCK_FIELDS.forEach((k) => {
      summedStock[k] = shedsWithNames.reduce((sum, s: any) => sum + Number(s[k] || 0), 0);
    });

    const payload: any = {
      ...form,
      company_id: Number(form.company_id),
      agent_id: Number(form.agent_id),
      farm_category: firstShed.farm_category,
      latitude:
        firstShed.latitude === "" || firstShed.latitude == null
          ? null
          : Number(firstShed.latitude),
      longitude:
        firstShed.longitude === "" || firstShed.longitude == null
          ? null
          : Number(firstShed.longitude),
      feed_type: firstShed.feed_type || "",
      ...summedStock,
    };

    let savedFarmer: any = null;
    try {
      if (editingId) {
        savedFarmer = await updateFarmer(editingId, payload);
      } else {
        savedFarmer = await createFarmer(payload);
      }
    } catch (err: any) {
      alert(err?.message || "Failed to save farmer.");
      return;
    }

    // Resolve the farmer id (farmerService may or may not return the row).
    let farmerId: number | null =
      editingId ||
      savedFarmer?.id ||
      savedFarmer?.data?.id ||
      savedFarmer?.[0]?.id ||
      null;
    if (!farmerId) {
      try {
        const { data } = await supabase
          .from("farmers")
          .select("id")
          .eq("farmer_name", form.farmer_name.trim())
          .eq("mobile", form.mobile || "")
          .order("created_at", { ascending: false })
          .limit(1);
        farmerId = data?.[0]?.id || null;
      } catch {
        farmerId = null;
      }
    }

    // Ownership (rule 2): stamp created_by on a new farmer, and on the
    // first edit of an old row that has none. Never overwrites an
    // existing creator. Best-effort: a missing created_by column must
    // not break the save.
    if (farmerId && currentUserId) {
      try {
        if (editingId) {
          const { data: existing } = await supabase
            .from("farmers")
            .select("created_by")
            .eq("id", farmerId)
            .maybeSingle();
          if (existing && !existing.created_by) {
            await supabase
              .from("farmers")
              .update({ created_by: currentUserId })
              .eq("id", farmerId);
          }
        } else {
          await supabase
            .from("farmers")
            .update({ created_by: currentUserId })
            .eq("id", farmerId);
        }
      } catch {
        // created_by column not added yet.
      }
    }

    // Save sheds (child table). Farmer save is NOT lost if this fails.
    if (farmerId) {
      try {
        const { data: existingSheds, error: existingErr } = await supabase
          .from("farmer_sheds")
          .select("id")
          .eq("farmer_id", farmerId);
        if (existingErr) throw existingErr;
        const keepIds = new Set(
          shedsWithNames
            .map((s: any) => s.id)
            .filter((id: any) => id != null && !String(id).startsWith("legacy-"))
            .map((id: any) => Number(id))
        );
        const staleIds = (existingSheds || [])
          .map((s: any) => Number(s.id))
          .filter((id: number) => !keepIds.has(id));
        if (staleIds.length > 0) {
          await supabase.from("farmer_sheds").delete().in("id", staleIds);
        }
        for (let i = 0; i < shedsWithNames.length; i++) {
          const s = shedsWithNames[i];
          const db = shedFormToDb(s, farmerId, form.farmer_name, i, shedsWithNames.length);
          const numericId =
            s.id != null && !String(s.id).startsWith("legacy-") ? Number(s.id) : null;
          if (numericId) {
            const { error } = await supabase
              .from("farmer_sheds")
              .update(db)
              .eq("id", numericId);
            if (error) throw error;
          } else {
            const { error } = await supabase.from("farmer_sheds").insert([db]);
            if (error) throw error;
          }
        }
        // Ownership: stamp sheds that still have no creator with the
        // farmer's creator (new farmer: current user; edit: the
        // farmer's existing creator, else current user). Never
        // overwrites a non-null created_by (rule 2).
        try {
          const farmerCreator =
            (editingId
              ? farmers.find((f: any) => Number(f.id) === Number(editingId))
                  ?.created_by
              : null) || currentUserId || null;
          if (farmerCreator) {
            await supabase
              .from("farmer_sheds")
              .update({ created_by: farmerCreator })
              .eq("farmer_id", farmerId)
              .is("created_by", null);
          }
        } catch {
          // created_by column not added yet on farmer_sheds.
        }
      } catch (shedErr: any) {
        alert(
          "Farmer saved, but sheds could not be saved. Please run farmer_sheds-migration.sql in Supabase first, then edit the farmer again to add sheds."
        );
      }
    }

    alert(editingId ? "Farmer Updated" : "Farmer Saved");
    setEditingId(null);
    setForm({ ...EMPTY_FORM });
    setFormSheds([{ ...EMPTY_SHED_FORM }]);
    setPickingShedIndex(null);
    setShowForm(false);
    loadData();
  }

  async function handleDelete(id: number) {
    const target = farmers.find((f: any) => Number(f.id) === Number(id));
    if (!canEditFarmer(target || {})) {
      alert("You can only delete farmers you added.");
      return;
    }
    if (!confirm("Delete Farmer? This will delete the farmer and all of their sheds.")) return;
    await deleteFarmer(id);
    setSelectedFarmer(null);
    loadData();
  }

 

  // Universal Location Helper (Works on both Android Native App & Browser)
  async function getDeviceCoordinates(): Promise<{ lat: number; lng: number }> {
    if (Capacitor.isNativePlatform()) {
      let perm = await Geolocation.checkPermissions();
      if (perm.location !== "granted") {
        perm = await Geolocation.requestPermissions();
      }
      if (perm.location !== "granted") {
        throw new Error(
          "Location permission denied. Please allow GPS access."
        );
      }

      try {
        const pos = await Geolocation.getCurrentPosition({
          enableHighAccuracy: true,
          timeout: 10000,
          maximumAge: 0,
        });
        return { lat: pos.coords.latitude, lng: pos.coords.longitude };
      } catch {
        const fallbackPos = await Geolocation.getCurrentPosition({
          enableHighAccuracy: false,
          timeout: 10000,
        });
        return {
          lat: fallbackPos.coords.latitude,
          lng: fallbackPos.coords.longitude,
        };
      }
    } else {
      return new Promise((resolve, reject) => {
        if (!navigator.geolocation) {
          reject(new Error("Geolocation is not supported by your browser."));
          return;
        }
        navigator.geolocation.getCurrentPosition(
          (pos) =>
            resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
          (err) => reject(err),
          { enableHighAccuracy: true, timeout: 10000 }
        );
      });
    }
  }

  async function handleLocateMeInPicker() {
    const targetIndex = pickingShedIndex ?? 0;
    setLocatingInPicker(true);
    try {
      const { lat, lng } = await getDeviceCoordinates();
      setPickerCenterOverride([lat, lng]);
      setFormSheds((prev) =>
        prev.map((s, i) =>
          i === targetIndex
            ? { ...s, latitude: lat.toFixed(6), longitude: lng.toFixed(6) }
            : s
        )
      );
    } catch (err: any) {
      alert(
        err?.message ||
          "Unable to retrieve your location. Please allow GPS access."
      );
    } finally {
      setLocatingInPicker(false);
    }
  }

  // ✅ Main Farmer Map-এর "Locate Me" বাটন হ্যান্ডলার (selectedFarmerPosition মুছে ফেলবে না, ফলে ডিসট্যান্স ও Polyline দেখাবে)
  async function handleLocateMeOnMainMap() {
    try {
      const { lat, lng } = await getDeviceCoordinates();
      setMyPosition([lat, lng]);
      setMyLocationMode(true);
    } catch (err: any) {
      alert(
        err?.message ||
          "Unable to retrieve your location. Please allow GPS access."
      );
    }
  }

  const filteredFarmers = farmers
    .filter((f) => {
      const q = search.toLowerCase();
      const matchSearch =
        !q ||
        f.farmer_name?.toLowerCase().includes(q) ||
        f.mobile?.toLowerCase().includes(q) ||
        f.village?.toLowerCase().includes(q) ||
        getShedsForFarmer(f).some((s: any, si: number) =>
          getShedDisplayName(f, s, si, getShedsForFarmer(f).length)
            .toLowerCase()
            .includes(q) ||
          getShedShortName(f, s, si, getShedsForFarmer(f).length).toLowerCase().includes(q)
        );
      const matchCompany =
        companyFilter === "" || String(f.company_id) === companyFilter;
      const matchAgent =
        agentFilter === "" || String(f.agent_id) === agentFilter;
      const matchDistrict =
        districtFilter === "" || f.district === districtFilter;
      const matchUpazila =
        upazilaFilter === "" || f.upazila === upazilaFilter;
      const matchUnion = unionFilter === "" || f.union_name === unionFilter;
      const matchCategory =
        farmCategoryFilter === "" ||
        getShedsForFarmer(f).some(
          (s: any) => s.farm_category === farmCategoryFilter
        );

      return (
        matchSearch &&
        matchCompany &&
        matchAgent &&
        matchDistrict &&
        matchUpazila &&
        matchUnion &&
        matchCategory
      );
    })
    .sort((a: any, b: any) => {
      const agentCompare = (a.agents?.agent_name || "").localeCompare(
        b.agents?.agent_name || ""
      );
      if (agentCompare !== 0) return agentCompare;
      return (a.farmer_name || "").localeCompare(b.farmer_name || "");
    });

  const agentGroupedFarmers = Object.values(
    filteredFarmers.reduce((acc: Record<string, any>, farmer: any) => {
      const key = String(farmer.agent_id || "unassigned");
      if (!acc[key]) {
        acc[key] = {
          agentId: key,
          agentName: farmer.agents?.agent_name || "Unassigned Agent",
          companyName: farmer.companies?.company_name || "-",
          upazila: farmer.upazila || "",
          farmers: [],
        };
      }
      acc[key].farmers.push(farmer);
      return acc;
    }, {})
  );

  const nearestFarmers = myPosition
    ? filteredFarmers
        .flatMap((f: any) =>
          getShedsForFarmer(f)
            .filter(
              (s: any) =>
                s.latitude &&
                s.longitude &&
                !isNaN(Number(s.latitude)) &&
                !isNaN(Number(s.longitude))
            )
            .map((s: any, sIdx: number) => {
              const km = calcDistanceKm(
                myPosition[0],
                myPosition[1],
                Number(s.latitude),
                Number(s.longitude)
              );
              return {
                ...f,
                shed: s,
                shedDisplayName: getShedDisplayName(
                  f,
                  s,
                  sIdx,
                  getShedsForFarmer(f).length
                ),
                latitude: s.latitude,
                longitude: s.longitude,
                distanceKm: km,
                distanceText: formatDistance(km),
              };
            })
        )
        .sort((a, b) => a.distanceKm - b.distanceKm)
        .slice(0, 10)
    : [];

  const activeDistanceText =
    myPosition && selectedFarmerPosition
      ? formatDistance(
          calcDistanceKm(
            myPosition[0],
            myPosition[1],
            selectedFarmerPosition[0],
            selectedFarmerPosition[1]
          )
        )
      : null;

  const companySummary = Object.values(
    filteredFarmers.reduce((acc: any, farmer: any) => {
      const company = farmer.companies?.company_name || "Unknown";
      if (!acc[company]) {
        acc[company] = {
          company,
          totalFarmers: 0,
          poultryFarmers: 0,
          cattleFarmers: 0,
          fishFarmers: 0,
          broiler: 0,
          layer: 0,
          sonali: 0,
          duck: 0,
          quail: 0,
          deshiChicken: 0,
          dairyCow: 0,
          bull: 0,
          calfHeifer: 0,
          pangus: 0,
          tilapia: 0,
          koi: 0,
          shing: 0,
          pabda: 0,
          gulsha: 0,
          magur: 0,
          carp: 0,
          golda: 0,
          bagda: 0,
          tengra: 0,
          totalPoultry: 0,
          totalCattle: 0,
          totalFish: 0,
          broilerFarmers: 0,
          layerFarmers: 0,
          sonaliFarmers: 0,
          duckFarmers: 0,
          quailFarmers: 0,
          deshiFarmers: 0,
          dairyCowFarmers: 0,
          bullFarmers: 0,
          calfHeiferFarmers: 0,
          pangusFarmers: 0,
          tilapiaFarmers: 0,
          koiFarmers: 0,
          shingFarmers: 0,
          pabdaFarmers: 0,
          gulshaFarmers: 0,
          magurFarmers: 0,
          carpFarmers: 0,
          goldaFarmers: 0,
          bagdaFarmers: 0,
          tengraFarmers: 0,
        };
      }

      const c = acc[company];
      c.totalFarmers++;
      const shedFieldCounts: Record<string, string> = {
        broiler: "broilerFarmers", layer: "layerFarmers", sonali: "sonaliFarmers",
        duck: "duckFarmers", quail: "quailFarmers", deshi_chicken: "deshiFarmers",
        dairy_cow: "dairyCowFarmers", bull: "bullFarmers", calf_heifer: "calfHeiferFarmers",
      };
      getShedsForFarmer(farmer).forEach((shed: any) => {
        [...POULTRY_FIELDS, ...CATTLE_FIELDS].forEach((k) => {
          if (Number(shed[k] || 0) > 0) c[shedFieldCounts[k]]++;
        });
        if (shed.farm_category === "Poultry") {
          c.poultryFarmers++;
          c.broiler += Number(shed.broiler || 0);
          c.layer += Number(shed.layer || 0);
          c.sonali += Number(shed.sonali || 0);
          c.duck += Number(shed.duck || 0);
          c.quail += Number(shed.quail || 0);
          c.deshiChicken += Number(shed.deshi_chicken || 0);
          c.totalPoultry =
            c.broiler + c.layer + c.sonali + c.duck + c.quail + c.deshiChicken;
        }
        if (shed.farm_category === "Cattle") {
          c.cattleFarmers++;
          c.dairyCow += Number(shed.dairy_cow || 0);
          c.bull += Number(shed.bull || 0);
          c.calfHeifer += Number(shed.calf_heifer || 0);
          c.totalCattle = c.dairyCow + c.bull + c.calfHeifer;
        }
        if (shed.farm_category === "Fish") {
          c.fishFarmers++;
          FISH_FIELDS.forEach((f) => {
            const val = Number(shed[f] || 0);
            c[f] += val;
            if (val > 0) c[`${f}Farmers`]++;
          });
          c.totalFish = FISH_FIELDS.reduce((sum, f) => sum + c[f], 0);
        }
      });
      return acc;
    }, {})
  );

  const selectedCompanyName =
    companies.find((c) => String(c.id) === companyFilter)?.company_name ||
    "All Companies";

  function styleStandardWorksheet(
    worksheet: ExcelJS.Worksheet,
    headerColor: string,
    widths: number[]
  ) {
    worksheet.getRow(1).eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: headerColor },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFFFFF" } },
        left: { style: "thin", color: { argb: "FFFFFF" } },
        bottom: { style: "thin", color: { argb: "FFFFFF" } },
        right: { style: "thin", color: { argb: "FFFFFF" } },
      };
    });

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin", color: { argb: "D1D5DB" } },
            left: { style: "thin", color: { argb: "D1D5DB" } },
            bottom: { style: "thin", color: { argb: "D1D5DB" } },
            right: { style: "thin", color: { argb: "D1D5DB" } },
          };
          cell.alignment = { vertical: "middle", wrapText: true };
        });
      }
    });

    widths.forEach((w, i) => {
      worksheet.getColumn(i + 1).width = w;
    });
    worksheet.getRow(1).height = 28;
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.pageSetup.orientation = "landscape";
    worksheet.pageSetup.fitToPage = true;
    worksheet.pageSetup.fitToWidth = 1;
    worksheet.pageSetup.fitToHeight = 0;
  }

  async function saveWorkbook(workbook: ExcelJS.Workbook, fileName: string) {
    const buffer = await workbook.xlsx.writeBuffer();
    await saveBufferOrBlob(buffer, fileName);
  }

  async function saveBufferOrBlob(
    bufferOrBlob: ArrayBuffer | Blob,
    fileName: string
  ) {
    try {
      const arrayBuffer =
        bufferOrBlob instanceof Blob
          ? await bufferOrBlob.arrayBuffer()
          : bufferOrBlob;

      if (Capacitor.isNativePlatform()) {
        let binary = "";
        const bytes = new Uint8Array(arrayBuffer);
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
          binary += String.fromCharCode.apply(
            null,
            Array.from(bytes.subarray(i, i + chunkSize))
          );
        }
        const base64Data = window.btoa(binary);

        const savedFile = await Filesystem.writeFile({
          path: fileName,
          data: base64Data,
          directory: Directory.Documents,
          recursive: true,
        });

        alert(`File saved successfully:\n📁 Documents/${fileName}`);

        if (savedFile.uri) {
          try {
            await Share.share({
              title: fileName,
              url: savedFile.uri,
              dialogTitle: "Open or View Excel File",
            });
          } catch {
            // user cancelled share dialog, ignore
          }
        }
      } else {
        const blob =
          bufferOrBlob instanceof Blob
            ? bufferOrBlob
            : new Blob([arrayBuffer], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
      }
    } catch (err: any) {
      console.error("Export save error:", err);
      alert("Export save error: " + (err?.message || ""));
    }
  }

  async function exportAllFarmersExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("All Farmers");
      worksheet.views = [{ state: "frozen", ySplit: 2 }];

      const generalHeaders = [
        "Company",
        "Agent",
        "Farmer",
        "Shed",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
      ];
      generalHeaders.forEach((h, idx) => {
        const col = idx + 1;
        worksheet.mergeCells(1, col, 2, col);
        worksheet.getCell(1, col).value = h;
      });

      worksheet.mergeCells("K1:P1");
      worksheet.getCell("K1").value = "POULTRY";
      ["Broiler", "Layer", "Sonali", "Duck", "Quail", "Deshi Chicken"].forEach(
        (v, i) => (worksheet.getCell(2, 11 + i).value = v)
      );

      worksheet.mergeCells("Q1:S1");
      worksheet.getCell("Q1").value = "CATTLE";
      ["Dairy Cow", "Bull", "Calf / Heifer"].forEach(
        (v, i) => (worksheet.getCell(2, 17 + i).value = v)
      );

      worksheet.mergeCells("T1:AD1");
      worksheet.getCell("T1").value = "FISH";
      [
        "Pangus",
        "Tilapia",
        "Koi",
        "Shing",
        "Pabda",
        "Gulsha",
        "Magur",
        "Carp",
        "Golda",
        "Bagda",
        "Tengra",
      ].forEach((v, i) => (worksheet.getCell(2, 20 + i).value = v));

      worksheet.mergeCells("AE1:AE2");
      worksheet.getCell("AE1").value = "Feed Type";
      worksheet.mergeCells("AF1:AF2");
      worksheet.getCell("AF1").value = "Status";

      for (let col = 1; col <= 32; col++) {
        if (col <= 10 || col >= 31) {
          const c = worksheet.getCell(1, col);
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "1E3A8A" },
          };
          c.font = { bold: true, color: { argb: "FFFFFFFF" } };
          c.alignment = {
            horizontal: "center",
            vertical: "middle",
            wrapText: true,
          };
        }
      }

      const styleGroup = (
        mainCell: string,
        mainBg: string,
        startCol: number,
        endCol: number,
        subBg: string,
        subFg: string
      ) => {
        const mc = worksheet.getCell(mainCell);
        mc.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: mainBg },
        };
        mc.font = { bold: true, color: { argb: "FFFFFFFF" } };
        mc.alignment = { horizontal: "center", vertical: "middle" };
        for (let col = startCol; col <= endCol; col++) {
          const sc = worksheet.getCell(2, col);
          sc.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: subBg },
          };
          sc.font = { bold: true, color: { argb: subFg } };
          sc.alignment = {
            horizontal: "center",
            vertical: "middle",
            wrapText: true,
          };
        }
      };

      styleGroup("K1", "D97706", 11, 16, "FEF3C7", "92400E");
      styleGroup("Q1", "15803D", 17, 19, "DCFCE7", "166534");
      styleGroup("T1", "0284C7", 20, 30, "E0F2FE", "075985");

      for (let r = 1; r <= 2; r++) {
        for (let c = 1; c <= 32; c++) {
          worksheet.getCell(r, c).border = {
            top: { style: "thin", color: { argb: "CBD5E1" } },
            left: { style: "thin", color: { argb: "CBD5E1" } },
            bottom: { style: "thin", color: { argb: "CBD5E1" } },
            right: { style: "thin", color: { argb: "CBD5E1" } },
          };
        }
      }

      filteredFarmers
        .flatMap((f: any) => {
          const sheds = getShedsForFarmer(f);
          return sheds.map((shed: any, sIdx: number) => ({
            f,
            shed,
            sIdx,
            totalSheds: sheds.length,
          }));
        })
        .forEach(({ f, shed, sIdx, totalSheds }: any) => {
        const row = worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          getShedDisplayName(f, shed, sIdx, totalSheds),
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          shed.farm_category || "",
          ...POULTRY_FIELDS.map((k) => Number(shed[k] || 0)),
          ...CATTLE_FIELDS.map((k) => Number(shed[k] || 0)),
          ...FISH_FIELDS.map((k) => Number(shed[k] || 0)),
          shed.feed_type || "",
          f.status || "",
        ]);
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin", color: { argb: "D1D5DB" } },
            left: { style: "thin", color: { argb: "D1D5DB" } },
            bottom: { style: "thin", color: { argb: "D1D5DB" } },
            right: { style: "thin", color: { argb: "D1D5DB" } },
          };
          cell.alignment = { vertical: "middle", wrapText: true };
        });
      });

      worksheet.getRow(1).height = 26;
      worksheet.getRow(2).height = 30;
      const widths = [
        22, 20, 22, 20, 15, 18, 18, 18, 18, 18, 12, 12, 12, 12, 12, 16, 14, 12,
        16, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 18, 12,
      ];
      widths.forEach((w, i) => (worksheet.getColumn(i + 1).width = w));
      worksheet.autoFilter = { from: "A2", to: "AF2" };

      await saveWorkbook(
        workbook,
        `All_Farmers_Report_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch (err: any) {
      console.error("All farmers export error:", err);
      alert("Excel export failed.");
    }
  }

  async function exportPoultryExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Poultry Farmers");
      const list = filteredFarmers.flatMap((f: any) => {
        const sheds = getShedsForFarmer(f);
        return sheds
          .map((shed: any, sIdx: number) => ({
            f,
            shed,
            sIdx,
            totalSheds: sheds.length,
          }))
          .filter((x: any) => x.shed.farm_category === "Poultry");
      });

      worksheet.addRow([
        "Company",
        "Agent",
        "Farmer",
        "Shed",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
        "Broiler",
        "Layer",
        "Sonali",
        "Duck",
        "Quail",
        "Deshi Chicken",
        "Feed Type",
        "Latitude",
        "Longitude",
        "Status",
        "Remarks",
      ]);

      list.forEach(({ f, shed, sIdx, totalSheds }: any) => {
        worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          getShedDisplayName(f, shed, sIdx, totalSheds),
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          shed.farm_category || "",
          ...POULTRY_FIELDS.map((k) => Number(shed[k] || 0)),
          shed.feed_type || "",
          shed.latitude ?? "",
          shed.longitude ?? "",
          f.status || "",
          f.remarks || "",
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "D97706",
        [
          25, 22, 25, 20, 16, 20, 20, 18, 18, 16, 12, 12, 12, 12, 12, 18, 18,
          14, 14, 12, 30,
        ]
      );
      await saveWorkbook(
        workbook,
        `Poultry_Farmers_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch (err: any) {
      console.error("Poultry export error:", err);
      alert("Poultry export failed.");
    }
  }

  async function exportCattleExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Cattle Farmers");
      const list = filteredFarmers.flatMap((f: any) => {
        const sheds = getShedsForFarmer(f);
        return sheds
          .map((shed: any, sIdx: number) => ({
            f,
            shed,
            sIdx,
            totalSheds: sheds.length,
          }))
          .filter((x: any) => x.shed.farm_category === "Cattle");
      });

      worksheet.addRow([
        "Company",
        "Agent",
        "Farmer",
        "Shed",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
        "Dairy Cow",
        "Bull",
        "Calf / Heifer",
        "Total Cattle",
        "Feed Type",
        "Latitude",
        "Longitude",
        "Status",
        "Remarks",
      ]);

      list.forEach(({ f, shed, sIdx, totalSheds }: any) => {
        const dc = Number(shed.dairy_cow || 0);
        const bl = Number(shed.bull || 0);
        const ch = Number(shed.calf_heifer || 0);
        worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          getShedDisplayName(f, shed, sIdx, totalSheds),
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          shed.farm_category || "",
          dc,
          bl,
          ch,
          dc + bl + ch,
          shed.feed_type || "",
          shed.latitude ?? "",
          shed.longitude ?? "",
          f.status || "",
          f.remarks || "",
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "15803D",
        [
          25, 22, 25, 20, 16, 20, 20, 18, 18, 16, 15, 12, 18, 18, 18, 14, 14,
          12, 30,
        ]
      );
      await saveWorkbook(
        workbook,
        `Cattle_Farmers_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch (err: any) {
      console.error("Cattle export error:", err);
      alert("Cattle export failed.");
    }
  }

  async function exportFishExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Fish Farmers");
      const list = filteredFarmers.flatMap((f: any) => {
        const sheds = getShedsForFarmer(f);
        return sheds
          .map((shed: any, sIdx: number) => ({
            f,
            shed,
            sIdx,
            totalSheds: sheds.length,
          }))
          .filter((x: any) => x.shed.farm_category === "Fish");
      });

      worksheet.addRow([
        "Company",
        "Agent",
        "Farmer",
        "Shed",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
        "Pangus",
        "Tilapia",
        "Koi",
        "Shing",
        "Pabda",
        "Gulsha",
        "Magur",
        "Carp",
        "Golda",
        "Bagda",
        "Tengra",
        "Total Fish",
        "Feed Type",
        "Latitude",
        "Longitude",
        "Status",
        "Remarks",
      ]);

      list.forEach(({ f, shed, sIdx, totalSheds }: any) => {
        const vals = FISH_FIELDS.map((k) => Number(shed[k] || 0));
        const total = vals.reduce((a, b) => a + b, 0);
        worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          getShedDisplayName(f, shed, sIdx, totalSheds),
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          shed.farm_category || "",
          ...vals,
          total,
          shed.feed_type || "",
          shed.latitude ?? "",
          shed.longitude ?? "",
          f.status || "",
          f.remarks || "",
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "0284C7",
        [
          25, 22, 25, 20, 16, 20, 20, 18, 18, 16, 12, 12, 12, 12, 12, 12, 12,
          12, 12, 12, 12, 15, 18, 14, 14, 12, 30,
        ]
      );
      await saveWorkbook(
        workbook,
        `Fish_Farmers_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch (err: any) {
      console.error("Fish export error:", err);
      alert("Fish export failed.");
    }
  }

  async function exportFarmerSummaryExcel() {
    try {
      const exportData = companySummary.map((row: any) => ({
        Company: row.company,
        Total_Farmers: row.totalFarmers,
        Poultry_Farmers: row.poultryFarmers,
        Cattle_Farmers: row.cattleFarmers,
        Fish_Farmers: row.fishFarmers,
      }));
      const worksheet = XLSX.utils.json_to_sheet(exportData);
      worksheet["!cols"] = [
        { wch: 28 },
        { wch: 18 },
        { wch: 20 },
        { wch: 20 },
        { wch: 18 },
      ];
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Company Summary");

      const today = new Date().toISOString().split("T")[0];
      const fileName = `Company_Farmer_Summary_${today}.xlsx`;

      const excelBuffer = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
      });

      await saveBufferOrBlob(excelBuffer, fileName);
    } catch (err: any) {
      console.error("Summary export error:", err);
      alert("Company Summary export failed.");
    }
  }

  const exportPoultrySummaryExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Poultry Summary");
      worksheet.addRow([
        "Company",
        "Broiler Farms",
        "Total Broiler",
        "Layer Farms",
        "Total Layer",
        "Sonali Farms",
        "Total Sonali",
        "Duck Farms",
        "Total Duck",
        "Quail Farms",
        "Total Quail",
        "Deshi Chicken Farms",
        "Total Deshi Chicken",
        "Total Poultry Farms",
        "Total Poultry",
      ]);

      companySummary.forEach((row: any) => {
        worksheet.addRow([
          row.company,
          row.broilerFarmers,
          row.broiler,
          row.layerFarmers,
          row.layer,
          row.sonaliFarmers,
          row.sonali,
          row.duckFarmers,
          row.duck,
          row.quailFarmers,
          row.quail,
          row.deshiFarmers,
          row.deshiChicken,
          row.poultryFarmers,
          row.totalPoultry,
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "D97706",
        [25, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18]
      );
      await saveWorkbook(
        workbook,
        `Poultry_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Poultry summary export failed.");
    }
  };

  const exportCattleSummaryExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Cattle Summary");
      worksheet.addRow([
        "Company",
        "Dairy Farms",
        "Total Dairy Cow",
        "Bull Farms",
        "Total Bull",
        "Calf / Heifer Farms",
        "Total Calf / Heifer",
        "Total Cattle",
      ]);
      companySummary.forEach((row: any) => {
        worksheet.addRow([
          row.company,
          row.dairyCowFarmers,
          row.dairyCow,
          row.bullFarmers,
          row.bull,
          row.calfHeiferFarmers,
          row.calfHeifer,
          row.totalCattle,
        ]);
      });
      styleStandardWorksheet(
        worksheet,
        "15803D",
        [28, 16, 16, 16, 16, 18, 18, 18]
      );
      await saveWorkbook(
        workbook,
        `Cattle_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Cattle Summary export failed.");
    }
  };

  const exportFishSummaryExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Fish Summary");
      worksheet.addRow([
        "Company",
        "Pangus Farms",
        "Total Pangus",
        "Tilapia Farms",
        "Total Tilapia",
        "Koi Farms",
        "Total Koi",
        "Shing Farms",
        "Total Shing",
        "Pabda Farms",
        "Total Pabda",
        "Gulsha Farms",
        "Total Gulsha",
        "Magur Farms",
        "Total Magur",
        "Carp Farms",
        "Total Carp",
        "Golda Farms",
        "Total Golda",
        "Bagda Farms",
        "Total Bagda",
        "Tengra Farms",
        "Total Tengra",
        "Total Fish Farms",
        "Total Fish",
      ]);

      companySummary.forEach((row: any) => {
        worksheet.addRow([
          row.company,
          row.pangusFarmers,
          row.pangus,
          row.tilapiaFarmers,
          row.tilapia,
          row.koiFarmers,
          row.koi,
          row.shingFarmers,
          row.shing,
          row.pabdaFarmers,
          row.pabda,
          row.gulshaFarmers,
          row.gulsha,
          row.magurFarmers,
          row.magur,
          row.carpFarmers,
          row.carp,
          row.goldaFarmers,
          row.golda,
          row.bagdaFarmers,
          row.bagda,
          row.tengraFarmers,
          row.tengra,
          row.fishFarmers,
          row.totalFish,
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "0284C7",
        [25, ...Array(24).fill(16)]
      );
      await saveWorkbook(
        workbook,
        `Fish_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Fish summary export failed.");
    }
  };

  function openEditForm(farmer: any) {
    if (!canEditFarmer(farmer)) {
      alert("You can only edit farmers you added.");
      return;
    }
    setEditingId(farmer.id);
    setForm({
      company_id: String(farmer.company_id),
      agent_id: String(farmer.agent_id),
      farmer_name: farmer.farmer_name || "",
      mobile: farmer.mobile || "",
      village: farmer.village || "",
      union_name: farmer.union_name || "",
      upazila: farmer.upazila || "",
      district: farmer.district || "",
      farm_category: farmer.farm_category || "",
      broiler: String(farmer.broiler || ""),
      layer: String(farmer.layer || ""),
      sonali: String(farmer.sonali || ""),
      duck: String(farmer.duck || ""),
      quail: String(farmer.quail || ""),
      deshi_chicken: String(farmer.deshi_chicken || ""),
      dairy_cow: String(farmer.dairy_cow || ""),
      bull: String(farmer.bull || ""),
      calf_heifer: String(farmer.calf_heifer || ""),
      pangus: String(farmer.pangus || ""),
      tilapia: String(farmer.tilapia || ""),
      koi: String(farmer.koi || ""),
      shing: String(farmer.shing || ""),
      pabda: String(farmer.pabda || ""),
      gulsha: String(farmer.gulsha || ""),
      magur: String(farmer.magur || ""),
      carp: String(farmer.carp || ""),
      golda: String(farmer.golda || ""),
      bagda: String(farmer.bagda || ""),
      tengra: String(farmer.tengra || ""),
      feed_type: farmer.feed_type || "",
      latitude: String(farmer.latitude || ""),
      longitude: String(farmer.longitude || ""),
      status: farmer.status || "Active",
      remarks: farmer.remarks || "",
    });
    setFormSheds(
      getShedsForFarmer(farmer).map((shed: any, sIdx: number) => {
        const row: any = {
          id: String(shed.id).startsWith("legacy-") ? undefined : shed.id,
          shed_name: getShedShortName(farmer, shed, sIdx, getShedsForFarmer(farmer).length),
          farm_category: shed.farm_category || "",
          feed_type: shed.feed_type || "",
          latitude: shed.latitude != null ? String(shed.latitude) : "",
          longitude: shed.longitude != null ? String(shed.longitude) : "",
          status: shed.status || "Active",
        };
        STOCK_FIELDS.forEach((k) => { row[k] = String(shed[k] || ""); });
        return row;
      })
    );
    setPickingShedIndex(null);
    setShowForm(true);
  }

  return (
    <div className="p-2.5 sm:p-4 max-w-7xl mx-auto space-y-3">
      {/* 1. COMPACT FILTER BAR */}
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2 items-center">
          <div className="col-span-2 sm:col-span-2 lg:col-span-2 relative">
            <Search
              size={13}
              className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Search farmer, shed, mobile, village..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full rounded-lg border border-slate-300 pl-8 pr-2.5 py-1.5 text-xs outline-none focus:border-blue-500"
            />
          </div>

          <SearchableSelect
            searchable={false}
            value={districtFilter}
            onChange={(v) => {
              setDistrictFilter(v);
              setUpazilaFilter("");
              setUnionFilter("");
              setAgentFilter("");
            }}
            placeholder="All Districts"
            options={districts.map((d) => ({ value: d.name, label: d.name }))}
          />

          <SearchableSelect
            value={upazilaFilter}
            onChange={(v) => {
              setUpazilaFilter(v);
              setUnionFilter("");
              setAgentFilter("");
            }}
            placeholder="All Upazilas"
            options={filterUpazilas.map((u: any) => ({
              value: u.name,
              label: u.name,
            }))}
          />

          <SearchableSelect
            value={unionFilter}
            onChange={setUnionFilter}
            placeholder="All Unions"
            options={filterUnions.map((u: any) => ({
              value: u.name,
              label: u.name,
            }))}
          />

          <SearchableSelect
            searchable={false}
            value={farmCategoryFilter}
            onChange={setFarmCategoryFilter}
            placeholder="All Categories"
            options={[
              { value: "Poultry", label: "🐔 Poultry" },
              { value: "Cattle", label: "🐄 Cattle" },
              { value: "Fish", label: "🐟 Fish" },
            ]}
          />

          <SearchableSelect
            value={companyFilter}
            onChange={(v) => {
              setCompanyFilter(v);
              setAgentFilter("");
            }}
            placeholder="All Companies"
            options={companies.map((c) => ({
              value: String(c.id),
              label: c.company_name,
            }))}
          />

          <div className="flex items-center gap-1.5">
            <SearchableSelect
              className="w-full"
              value={agentFilter}
              onChange={setAgentFilter}
              placeholder="All Agents"
              options={availableFilterAgents.map((a) => ({
                value: String(a.id),
                label:
                  a.agent_name +
                  (isAgentNonExclusive(a) ? " (non-ex.)" : ""),
              }))}
            />

            <button
              onClick={clearFilters}
              title="Reset Filters"
              className="shrink-0 inline-flex items-center justify-center rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2.5 py-1.5 text-xs font-semibold text-slate-700 transition"
            >
              <RotateCcw size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* 2. ACTION BAR */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-3.5 py-2 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
          {canAddFarmerHere && (
          <button
            onClick={() => {
              setEditingId(null);
              const nourishComp =
                companies.find((c) =>
                  c.company_name?.toLowerCase().includes("nourish")
                ) || companies[0];
              const defCompId = nourishComp ? String(nourishComp.id) : "";

              const sortedDists = [...districts].sort((a, b) =>
                (a.name || "").localeCompare(b.name || "")
              );
              const defDist = sortedDists.length > 0 ? sortedDists[0].name : "";
              const defUpa =
                !isFullAccess && assignedUpazilas.length > 0
                  ? [...assignedUpazilas].sort()[0]
                  : "";

              setForm({
                ...EMPTY_FORM,
                company_id: defCompId,
                district: defDist,
                upazila: defUpa,
              });
              setFormSheds([{ ...EMPTY_SHED_FORM }]);
              setPickingShedIndex(null);
              setShowForm(true);
            }}
            className="inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <Plus size={14} /> Add Farmer
          </button>
          )}

          <button
            onClick={() => {
              setShowFarmerMap(false);
              setShowSummary(true);
            }}
            className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <BarChart3 size={14} /> Summary
          </button>

          <div className="relative">
            <button
              onClick={() => setShowExportMenu(!showExportMenu)}
              className="inline-flex items-center gap-1.5 bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
            >
              <FileSpreadsheet size={14} /> Export Excel
              <span className="text-[10px]">{showExportMenu ? "▲" : "▼"}</span>
            </button>

            {showExportMenu && (
              <div className="absolute left-0 mt-1.5 w-48 bg-white border border-slate-200 rounded-xl shadow-xl z-[2000] text-xs overflow-hidden">
                <button
                  onClick={() => {
                    exportAllFarmersExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-slate-50 border-b font-medium"
                >
                  📄 Export All Farmers
                </button>
                <button
                  onClick={() => {
                    exportPoultryExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-amber-50 border-b font-medium"
                >
                  🐔 Export Poultry Farmers
                </button>
                <button
                  onClick={() => {
                    exportCattleExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-emerald-50 border-b font-medium"
                >
                  🐄 Export Cattle Farmers
                </button>
                <button
                  onClick={() => {
                    exportFishExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-sky-50 font-medium"
                >
                  🐟 Export Fish Farmers
                </button>
              </div>
            )}
          </div>

          <button
            onClick={() => {
              if (!showFarmerMap) {
                setMyLocationMode(false);
                setMyPosition(null);
                setSelectedFarmerPosition(null);
              }
              setShowSummary(false);
              setShowFarmerMap(!showFarmerMap);
            }}
            className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition ${
              showFarmerMap
                ? "bg-red-600 hover:bg-red-700 text-white"
                : "bg-blue-600 hover:bg-blue-700 text-white"
            }`}
          >
            <MapIcon size={14} />
            {showFarmerMap ? "Close Map" : "View Map"}
          </button>
        </div>

        <div className="text-xs font-semibold text-slate-600">
          Total Farmers:{" "}
          <span className="text-slate-900">{filteredFarmers.length}</span>
          {"  •  "}Total Sheds:{" "}
          <span className="text-slate-900">{getTotalSheds(filteredFarmers)}</span>
        </div>
      </div>

      {/* 3. INTERACTIVE MAP VIEW WITH FLOATING CLOSE BUTTON */}
      {showFarmerMap && (
        <div
          ref={mapSectionRef}
          className="bg-white rounded-xl shadow border border-slate-300 p-3 relative z-10"
        >
          <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
            <h3 className="text-sm font-bold text-slate-800">Farmer Map</h3>
            <div className="flex flex-wrap items-center gap-2 text-[11px] font-semibold">
              <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800">
                🐔 Broiler
              </span>
              <span className="px-2 py-0.5 rounded bg-orange-100 text-orange-800">
                🥚 Layer
              </span>
              <span className="px-2 py-0.5 rounded bg-red-100 text-red-800">
                🐓 Sonali
              </span>
              <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                🐄 Cattle
              </span>
              <span className="px-2 py-0.5 rounded bg-sky-100 text-sky-800">
                🐟 Fish
              </span>
            </div>
          </div>

          <div className="w-full overflow-hidden rounded-xl border border-slate-300 relative">
            <MapContainer
              center={[24.7471, 90.4203]}
              zoom={8}
              style={{ height: "480px", width: "100%", zIndex: 1 }}
            >
              <FitBounds
                skip={Boolean(selectedFarmerPosition || myPosition)}
                locations={filteredFarmers.flatMap((f: any) =>
                  getShedsForFarmer(f)
                    .filter((s: any) => s.latitude && s.longitude)
                    .map((s: any) => ({
                      lat: Number(s.latitude),
                      lng: Number(s.longitude),
                    }))
                )}
              />
              <MapLayers />

              {/* অটো ফোকাস ও বাউন্ডস কন্ট্রোলার */}
              <MapFocusController
                farmerPos={selectedFarmerPosition}
                myPos={myPosition}
                myLocationMode={myLocationMode}
              />

              {myPosition && (
                <Marker
                  position={myPosition}
                  ref={(ref) => {
                    myLocationMarkerRef.current = ref;
                  }}
                >
                  <Tooltip
                    permanent
                    direction="top"
                    offset={[0, -15]}
                    interactive={true}
                    eventHandlers={{
                      click: () => {
                        myLocationMarkerRef.current?.openPopup();
                      },
                    }}
                  >
                    <span className="cursor-pointer font-semibold">
                      📍 My Location
                    </span>
                  </Tooltip>
                  <Popup>
                    <div className="max-h-60 overflow-y-auto w-52 text-xs space-y-1.5">
                      <h3 className="font-bold text-slate-800">
                        📍 Nearest Farmers
                      </h3>
                      {nearestFarmers.length === 0 ? (
                        <div>No nearby farmers found.</div>
                      ) : (
                        nearestFarmers.map((farmer, index) => (
                          <div key={`${farmer.id}-${index}`} className="border-b pb-1">
                            <button
                              type="button"
                              onClick={() => {
                                const targetPos: [number, number] = [
                                  Number(farmer.latitude),
                                  Number(farmer.longitude),
                                ];
                                setMyLocationMode(false);
                                setSelectedFarmerPosition(targetPos);
                                setSelectedShed(farmer.shed || null);
                                setNewNoteText("");
                                loadFarmerNotes(farmer.id);
                                myLocationMarkerRef.current?.closePopup();
                              }}
                              className="text-blue-600 font-bold hover:underline text-left"
                            >
                              {index + 1}. {farmer.shedDisplayName || farmer.farmer_name}
                            </button>
                            <div className="text-slate-500">
                              {farmer.distanceText} away
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </Popup>
                </Marker>
              )}

              {/* My Location থেকে Selected Farmer পর্যন্ত ড্যাশ-ড্যাশ লাইন */}
              {myPosition && selectedFarmerPosition && (
                <Polyline
                  positions={[myPosition, selectedFarmerPosition]}
                  pathOptions={{
                    color: "#2563eb",
                    weight: 3,
                    dashArray: "8, 8",
                  }}
                />
              )}

              {filteredFarmers
                .flatMap((farmer: any) => {
                  const sheds = getShedsForFarmer(farmer);
                  return sheds
                    .map((shed: any, sIdx: number) => ({
                      farmer,
                      shed,
                      sIdx,
                      totalSheds: sheds.length,
                    }))
                    .filter((x: any) => x.shed.latitude && x.shed.longitude);
                })
                .map(({ farmer, shed, sIdx, totalSheds }: any) => {
                  const shedKey = shed.id
                    ? `shed-${shed.id}`
                    : `${farmer.id}-${sIdx}`;
                  const shedDisplayName = getShedDisplayName(
                    farmer,
                    shed,
                    sIdx,
                    totalSheds
                  );
                  const badge = getFarmerBadgeInfo(shed);
                  const breakdown = getShedStockBreakdown(shed);
                  const fLat = Number(shed.latitude);
                  const fLng = Number(shed.longitude);
                  const isHighlighted =
                    selectedFarmerPosition !== null &&
                    selectedFarmerPosition[0] === fLat &&
                    selectedFarmerPosition[1] === fLng;

                  const distFromMe =
                    myPosition && isHighlighted
                      ? formatDistance(
                          calcDistanceKm(
                            myPosition[0],
                            myPosition[1],
                            fLat,
                            fLng
                          )
                        )
                      : undefined;

                  return (
                    <Marker
                      key={shedKey}
                      ref={(ref) => {
                        farmerMarkerRefs.current[shedKey] = ref;
                      }}
                      position={[fLat, fLng]}
                      icon={buildCustomFarmMarker(
                        farmer,
                        shed,
                        isHighlighted,
                        distFromMe ? `${distFromMe} away` : undefined,
                        sIdx,
                        totalSheds
                      )}
                      eventHandlers={{
                        click: () => {
                          setSelectedFarmerPosition([fLat, fLng]);
                          setSelectedShed(shed);
                          setNewNoteText("");
                          loadFarmerNotes(farmer.id);
                        },
                      }}
                    >
                      <Popup maxWidth={340}>
                        <div className="w-[270px] sm:w-[290px] text-xs space-y-2">
                          <div className="flex items-center gap-2 border-b pb-1.5">
                            {farmer.companies?.logo_url && (
                              <img
                                src={farmer.companies?.logo_url}
                                className="w-9 h-9 object-contain border rounded p-0.5 bg-white shrink-0"
                              />
                            )}
                            <div>
                              <h3 className="font-bold text-sm text-slate-800 leading-tight">
                                {badge.emoji} {shedDisplayName}
                              </h3>
                              <p className="text-[11px] text-slate-500">
                                {farmer.farmer_name} •{" "}
                                {farmer.companies?.company_name} •{" "}
                                {farmer.agents?.agent_name}
                              </p>
                            </div>
                          </div>

                          <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-2 space-y-1.5">
                            <div className="font-bold text-[11px] text-emerald-900">
                              📝 Farmer Tracking
                            </div>
                            <div className="max-h-24 overflow-y-auto space-y-1 pr-0.5">
                              {loadingNotes &&
                              activeTrackingFarmerId === farmer.id ? (
                                <div className="text-[11px] text-slate-400">
                                  Loading...
                                </div>
                              ) : activeTrackingFarmerId === farmer.id &&
                                farmerNotes.length > 0 ? (
                                farmerNotes
                                  .slice()
                                  .reverse()
                                  .map((n) => (
                                    <div
                                      key={n.id}
                                      className="bg-white border border-emerald-100 rounded px-2 py-1 text-[11px] flex items-start gap-1.5"
                                    >
                                      <div className="mt-0.5 shrink-0">
                                        <UserAvatar
                                          name={n.user_name || "Officer"}
                                          avatarUrl={getUserAvatar(
                                            n.user_id,
                                            n.user_name
                                          )}
                                          size="w-4 h-4"
                                          textSize="text-[8px]"
                                        />
                                      </div>
                                      <div className="min-w-0 flex-1 break-words">
                                        <span className="font-bold text-emerald-800">
                                          {n.user_name}:
                                        </span>{" "}
                                        <span className="text-slate-700">
                                          {renderNoteWithMentions(n.note)}
                                        </span>
                                      </div>
                                    </div>
                                  ))
                              ) : (
                                <div className="text-[10px] text-slate-400">
                                  No tracking note yet.
                                </div>
                              )}
                            </div>

                            <form
                              onSubmit={(e) =>
                                handleAddFarmerNote(e, farmer.id)
                              }
                              className="flex gap-1"
                            >
                              <input
                                type="text"
                                placeholder="Write note..."
                                value={
                                  activeTrackingFarmerId === farmer.id
                                    ? newNoteText
                                    : ""
                                }
                                onChange={(e) => {
                                  setActiveTrackingFarmerId(farmer.id);
                                  setNewNoteText(e.target.value);
                                }}
                                className="flex-1 border border-emerald-300 rounded px-2 py-1 text-[11px] bg-white outline-none"
                              />
                              <button
                                type="submit"
                                disabled={savingNote}
                                className="bg-emerald-600 text-white px-2 py-1 rounded text-[11px] font-semibold"
                              >
                                Track
                              </button>
                            </form>
                          </div>

                          <table className="w-full border-collapse border border-slate-200 text-[11px]">
                            <tbody>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Mobile
                                </td>
                                <td className="border px-2 py-1">
                                  {farmer.mobile || "-"}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Address
                                </td>
                                <td className="border px-2 py-1">
                                  {[
                                    farmer.village,
                                    farmer.union_name,
                                    farmer.upazila,
                                    farmer.district,
                                  ]
                                    .filter(Boolean)
                                    .join(", ")}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Category
                                </td>
                                <td className="border px-2 py-1 font-bold">
                                  {shed.farm_category}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Quantities
                                </td>
                                <td className="border px-2 py-1 font-semibold text-blue-700">
                                  {breakdown.length === 0
                                    ? "-"
                                    : breakdown
                                        .map((b) => `${b.label}: ${b.qty}`)
                                        .join(", ")}
                                </td>
                              </tr>
                              {shed.feed_type && (
                                <tr>
                                  <td className="border px-2 py-1 font-semibold bg-slate-50">
                                    Feed
                                  </td>
                                  <td className="border px-2 py-1">
                                    {shed.feed_type}
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      </Popup>
                    </Marker>
                  );
                })}
            </MapContainer>

           
            {/* Floating Locate Me + Close Button (Middle Side-by-Side) */}
            <div className="absolute top-3 left-1/2 -translate-x-1/2 z-[500] flex flex-col items-center gap-1.5">
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={handleLocateMeOnMainMap}
                  className="flex items-center gap-1 bg-blue-600 hover:bg-blue-700 text-white rounded-full shadow-lg px-2.5 py-1.5 text-xs font-semibold whitespace-nowrap transition"
                >
                  <MapPin size={15} /> Locate Me
                </button>

                <button
                  type="button"
                  onClick={() => setShowFarmerMap(false)}
                  title="Close Map"
                  className="flex items-center gap-1 bg-purple-500 hover:bg-red-700 text-white rounded-full shadow-lg px-3 py-1.5 text-xs font-bold whitespace-nowrap transition"
                >
                  <X size={14} /> Close
                </button>
              </div>

              {activeDistanceText && (
                <div className="rounded-full bg-slate-900/90 text-white px-3 py-1 text-[11px] font-bold shadow-md whitespace-nowrap">
                  📍 Distance: {activeDistanceText}
                </div>
              )}
            </div>

            
          </div>
        </div>
      )}

      {/* 4. AGENT-WISE FARMER TABLE */}
      {agentGroupedFarmers.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-lg p-8 text-center text-xs text-slate-500">
          No farmers found for the selected filter.
        </div>
      ) : (
        <div className="space-y-4">
          {agentGroupedFarmers.map((group: any) => (
            <div
              key={group.agentId}
              className="bg-white border border-slate-300 rounded-lg overflow-hidden shadow-sm"
            >
              <div className="bg-slate-100 border-b border-slate-300 px-3 py-2 flex flex-wrap items-center justify-between gap-2">
                <div className="text-xs sm:text-sm font-bold text-slate-800">
                  Agent: {group.agentName}{" "}
                  <span className="font-normal text-slate-600">
                    ({group.companyName})
                  </span>
                </div>
                <span className="text-xs font-semibold bg-white border border-slate-300 px-2 py-0.5 rounded text-slate-700">
                  Farmers: {group.farmers.length}  •  Sheds:{" "}
                  {getTotalSheds(group.farmers)}
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-xs">
                  <thead className="bg-slate-50">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        SL
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left min-w-[180px] md:min-w-[200px] whitespace-nowrap">
                        Farmer
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Mobile
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Village
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Union
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Upazila
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Shed
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Category
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Farm Details
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Status
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">
                        Action
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {group.farmers.flatMap((farmer: any, idx: number) => {
                      const sheds = getShedsForFarmer(farmer);
                      const rowSpan = sheds.length;
                      const farmerVisits = (visitRows || []).filter(
                        (v: any) => String(v.farmer_id) === String(farmer.id)
                      );
                      return sheds.map((shed: any, sIdx: number) => {
                        const breakdown = getShedStockBreakdown(shed);
                        const shedShortName = getShedShortName(farmer, shed, sIdx, rowSpan);
                        // Status is PER SHED: a shed with a Running visit
                        // is Running; otherwise its own last sale decides
                        // (<=60 days Active, older Inactive); no history
                        // falls back to the stored shed.status.
                        const shedStatus = getShedVisitStatus(
                          farmer,
                          shed,
                          farmerVisits
                        ).status;
                        const shedKey = shed.id
                          ? `shed-${shed.id}`
                          : `${farmer.id}-${sIdx}`;
                        return (
                          <tr key={shedKey} className="hover:bg-slate-50">
                            {sIdx === 0 && (
                              <td
                                rowSpan={rowSpan}
                                className="border border-slate-300 px-2 py-1 align-top"
                              >
                                {idx + 1}
                              </td>
                            )}
                            {sIdx === 0 && (
                              <td
                                rowSpan={rowSpan}
                                className="border border-slate-300 px-2 py-1 font-semibold text-slate-800 align-top min-w-[180px] md:min-w-[200px] whitespace-nowrap"
                                title={farmer.farmer_name}
                              >
                                {farmer.farmer_name}
                                {rowSpan > 1 && (
                                  <span className="block text-[10px] font-semibold text-slate-500">
                                    {rowSpan} sheds
                                  </span>
                                )}
                              </td>
                            )}
                            {sIdx === 0 && (
                              <td
                                rowSpan={rowSpan}
                                className="border border-slate-300 px-2 py-1 align-top"
                              >
                                {farmer.mobile || "-"}
                              </td>
                            )}
                            {sIdx === 0 && (
                              <td
                                rowSpan={rowSpan}
                                className="border border-slate-300 px-2 py-1 align-top"
                              >
                                {farmer.village || "-"}
                              </td>
                            )}
                            {sIdx === 0 && (
                              <td
                                rowSpan={rowSpan}
                                className="border border-slate-300 px-2 py-1 align-top"
                              >
                                {farmer.union_name || "-"}
                              </td>
                            )}
                            {sIdx === 0 && (
                              <td
                                rowSpan={rowSpan}
                                className="border border-slate-300 px-2 py-1 align-top"
                              >
                                {farmer.upazila || "-"}
                              </td>
                            )}
                            <td
                              className="border border-slate-300 px-2 py-1 font-medium whitespace-nowrap truncate max-w-[110px]"
                              title={getShedDisplayName(farmer, shed, sIdx, rowSpan)}
                            >
                              {shedShortName}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 font-medium">
                              {shed.farm_category || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {breakdown.length === 0
                                ? "-"
                                : breakdown
                                    .map((b) => `${b.label}: ${b.qty}`)
                                    .join(", ")}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 whitespace-nowrap">
                              <span
                                className={`inline-block px-1.5 py-0 rounded-full text-[10px] font-bold ${
                                  shedStatus === "Running"
                                    ? "bg-emerald-100 text-emerald-800"
                                    : shedStatus === "Active"
                                    ? "bg-blue-100 text-blue-800"
                                    : "bg-amber-100 text-amber-800"
                                }`}
                              >
                                {shedStatus}
                              </span>
                            </td>
                            {/* Actions: NO rowSpan — every shed row has its own
                                View + Map buttons for that shed only. */}
                            <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap align-top">
                              <button
                                onClick={() => setSelectedFarmer(farmer)}
                                className="bg-blue-600 hover:bg-blue-700 text-white px-2 py-0.5 rounded text-[11px] font-semibold mr-1"
                              >
                                View
                              </button>
                              <button
                                onClick={() => handleViewFarmerOnMap(farmer, shed)}
                                className="bg-teal-600 hover:bg-teal-700 text-white px-2 py-0.5 rounded text-[11px] font-semibold inline-flex items-center gap-0.5"
                              >
                                <MapPin size={11} /> Map
                              </button>
                            </td>
                          </tr>
                        );
                      });
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* 5. VIEW FARMER MODAL */}
      {selectedFarmer && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl w-full max-w-xl max-h-[90vh] flex flex-col overflow-hidden shadow-2xl border border-slate-300">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-300 bg-slate-100">
              <h2 className="text-sm sm:text-base font-bold text-slate-800">
                Farmer: {selectedFarmer.farmer_name}
              </h2>
              <button
                onClick={() => setSelectedFarmer(null)}
                className="text-slate-500 hover:text-slate-800"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-4 space-y-3 overflow-y-auto">
              <div className="bg-emerald-50 border border-emerald-300 rounded-lg p-3 space-y-2">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-bold text-emerald-900">
                    📝 Farmer Tracking
                  </span>
                  <span className="text-emerald-700 flex items-center gap-1.5">
                    <span>User:</span>
                    <UserAvatar
                      name={currentUserName}
                      avatarUrl={getUserAvatar(undefined, currentUserName)}
                      size="w-4 h-4"
                      textSize="text-[8px]"
                    />
                    <b>{currentUserName}</b>
                  </span>
                </div>

                <div className="max-h-36 overflow-y-auto space-y-1.5">
                  {loadingNotes ? (
                    <div className="text-xs text-slate-500">Loading...</div>
                  ) : farmerNotes.length === 0 ? (
                    <div className="text-xs text-slate-400 bg-white border border-emerald-200 rounded p-2 text-center">
                      No tracking notes yet.
                    </div>
                  ) : (
                    farmerNotes
                      .slice()
                      .reverse()
                      .map((item) => (
                        <div
                          key={item.id}
                          className="bg-white border border-emerald-200 rounded px-2.5 py-1.5 text-xs flex justify-between items-start gap-2"
                        >
                          <div className="flex items-start gap-1.5 min-w-0 flex-1">
                            <div className="mt-0.5 shrink-0">
                              <UserAvatar
                                name={item.user_name || "Officer"}
                                avatarUrl={getUserAvatar(
                                  item.user_id,
                                  item.user_name
                                )}
                                size="w-4 h-4"
                                textSize="text-[8px]"
                              />
                            </div>
                            <div className="min-w-0 flex-1 break-words">
                              <span className="font-bold text-emerald-800">
                                {item.user_name}:
                              </span>{" "}
                              <span className="text-slate-800">
                                {renderNoteWithMentions(item.note)}
                              </span>
                            </div>
                          </div>
                          <span className="text-[10px] text-slate-400 shrink-0">
                            {new Date(item.created_at).toLocaleDateString(
                              "en-GB"
                            )}
                          </span>
                        </div>
                      ))
                  )}
                </div>

                <form
                  onSubmit={(e) => handleAddFarmerNote(e, selectedFarmer.id)}
                  className="flex gap-1.5"
                >
                  <input
                    type="text"
                    placeholder="Write tracking note..."
                    value={newNoteText}
                    onChange={(e) => setNewNoteText(e.target.value)}
                    className="flex-1 rounded border border-emerald-300 bg-white px-2.5 py-1.5 text-xs outline-none"
                  />
                  <button
                    type="submit"
                    disabled={savingNote || !newNoteText.trim()}
                    className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-3 py-1.5 rounded text-xs font-semibold flex items-center gap-1"
                  >
                    <Send size={12} /> Track
                  </button>
                </form>
              </div>

              <table className="w-full border-collapse border border-slate-300 text-xs">
                <tbody>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold w-1/3">
                      Company
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.companies?.company_name || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Agent
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.agents?.agent_name || "-"}
                      {(() => {
                        const ag = agents.find(
                          (a: any) =>
                            String(a.id) === String(selectedFarmer.agent_id)
                        );
                        const names = ag
                          ? getAgentLinkedCompanyNames(ag)
                          : [];
                        return names.length > 1 ? (
                          <span className="block text-[11px] font-normal text-slate-500">
                            Also: {names.join(", ")}
                          </span>
                        ) : null;
                      })()}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Farmer Name
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-bold">
                      {selectedFarmer.farmer_name}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Mobile
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.mobile || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Address
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {[
                        selectedFarmer.village,
                        selectedFarmer.union_name,
                        selectedFarmer.upazila,
                        selectedFarmer.district,
                      ]
                        .filter(Boolean)
                        .join(", ")}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Farm Category
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-semibold">
                      {getFarmerCategories(selectedFarmer).join(", ") ||
                        selectedFarmer.farm_category}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Stock / Capacity
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-semibold text-blue-700">
                      {getStockBreakdownList(selectedFarmer).length === 0
                        ? "-"
                        : getStockBreakdownList(selectedFarmer)
                            .map((b) => `${b.label}: ${b.qty}`)
                            .join(", ")}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Feed Type
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.feed_type || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Status
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.status || "Active"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Coordinates
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.latitude && selectedFarmer.longitude
                        ? `${selectedFarmer.latitude}, ${selectedFarmer.longitude}`
                        : "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Remarks
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.remarks || "-"}
                    </td>
                  </tr>
                </tbody>
              </table>

              <div>
                <h3 className="text-xs font-bold text-slate-700 mb-1.5">
                  Sheds ({getFarmerShedsCount(selectedFarmer)})
                </h3>
                <table className="w-full border-collapse border border-slate-300 text-xs">
                  <thead className="bg-slate-50">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1 text-left">Shed</th>
                      <th className="border border-slate-300 px-2 py-1 text-left">Category</th>
                      <th className="border border-slate-300 px-2 py-1 text-left">Status</th>
                      <th className="border border-slate-300 px-2 py-1 text-left">Location</th>
                      <th className="border border-slate-300 px-2 py-1 text-left">Stock</th>
                    </tr>
                  </thead>
                  <tbody>
                    {getShedsForFarmer(selectedFarmer).map((shed: any, sIdx: number) => {
                      const b = getShedStockBreakdown(shed);
                      const modalShedStatus = getShedVisitStatus(
                        selectedFarmer,
                        shed,
                        (visitRows || []).filter(
                          (v: any) =>
                            String(v.farmer_id) === String(selectedFarmer.id)
                        )
                      ).status;
                      return (
                        <tr key={shed.id || sIdx}>
                          <td className="border border-slate-300 px-2 py-1 font-semibold">
                            {getShedDisplayName(
                              selectedFarmer,
                              shed,
                              sIdx,
                              getShedsForFarmer(selectedFarmer).length
                            )}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {shed.farm_category || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 whitespace-nowrap">
                            <span
                              className={`inline-block px-1.5 py-0 rounded-full text-[10px] font-bold ${
                                modalShedStatus === "Running"
                                  ? "bg-emerald-100 text-emerald-800"
                                  : modalShedStatus === "Active"
                                  ? "bg-blue-100 text-blue-800"
                                  : "bg-amber-100 text-amber-800"
                              }`}
                            >
                              {modalShedStatus}
                            </span>
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {shed.latitude && shed.longitude
                              ? `${Number(shed.latitude).toFixed(5)}, ${Number(shed.longitude).toFixed(5)}`
                              : "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 font-semibold text-blue-700">
                            {b.length === 0 ? "-" : b.map((x) => `${x.label}: ${x.qty}`).join(", ")}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Whole-farmer delete lives ONLY here in the View modal. */}
              <div className="flex justify-end gap-2 pt-1">
                {canEditFarmer(selectedFarmer) && (
                  <>
                <button
                  onClick={() => {
                    openEditForm(selectedFarmer);
                    setSelectedFarmer(null);
                  }}
                  className="bg-yellow-500 hover:bg-yellow-600 text-white px-3 py-1.5 rounded text-xs font-semibold"
                >
                  Edit
                </button>
                <button
                  onClick={() => handleDelete(selectedFarmer.id)}
                  className="bg-red-600 hover:bg-red-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                >
                  Delete Farmer
                </button>
                  </>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 6. ADD / EDIT FARMER MODAL */}
      {showForm && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-5 w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h2 className="text-base font-bold text-slate-800">
                {editingId ? "Edit Farmer" : "Add Farmer"}
              </h2>
              <button onClick={() => setShowForm(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Company *
                </label>
                <SearchableSelect
                  value={form.company_id}
                  onChange={(companyId) =>
                    setForm({ ...form, company_id: companyId, agent_id: "" })
                  }
                  placeholder="Select Company"
                  options={companies.map((c) => ({
                    value: String(c.id),
                    label: c.company_name,
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  District *
                </label>
                <SearchableSelect
                  searchable={false}
                  value={form.district}
                  onChange={(v) =>
                    setForm({
                      ...form,
                      district: v,
                      upazila: "",
                      union_name: "",
                      agent_id: "",
                    })
                  }
                  placeholder="Select District"
                  options={districts.map((d) => ({
                    value: d.name,
                    label: d.name,
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Upazila *
                </label>
                <SearchableSelect
                  value={form.upazila}
                  onChange={(v) =>
                    setForm({
                      ...form,
                      upazila: v,
                      union_name: "",
                      agent_id: "",
                    })
                  }
                  placeholder="Select Upazila"
                  options={upazilas.map((u) => ({
                    value: u.name,
                    label: u.name,
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Agent *
                </label>
                <SearchableSelect
                  value={form.agent_id}
                  onChange={(v) => setForm({ ...form, agent_id: v })}
                  placeholder="Select Agent"
                  options={formAvailableAgents.map((a) => ({
                    value: String(a.id),
                    label:
                      a.agent_name +
                      (isAgentNonExclusive(a) ? " (non-ex.)" : ""),
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Farmer Name *
                </label>
                <input
                  placeholder="Farmer Name"
                  value={form.farmer_name}
                  onChange={(e) =>
                    setForm({ ...form, farmer_name: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Mobile
                </label>
                <input
                  placeholder="017XXXXXXXX"
                  value={form.mobile}
                  onChange={(e) => setForm({ ...form, mobile: e.target.value })}
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Union
                </label>
                <SearchableSelect
                  value={form.union_name}
                  onChange={(v) => setForm({ ...form, union_name: v })}
                  placeholder="Select Union"
                  options={unions.map((u) => ({
                    value: u.name,
                    label: u.name,
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Village
                </label>
                <input
                  placeholder="Village"
                  value={form.village}
                  onChange={(e) =>
                    setForm({ ...form, village: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>
            </div>

            {/* Sheds — one farmer, multiple sheds, each with its own location/category/stock */}
            <div className="mb-3">
              <div className="flex items-center justify-between mb-2">
                <label className="block text-xs font-bold text-slate-600">
                  Sheds * ({formSheds.length})
                </label>
                <span className="text-[11px] text-slate-500">
                  Each shed can be in a different place with its own category.
                </span>
              </div>

              {formSheds.map((shed: any, sIdx: number) => (
                <div
                  key={shed.id || sIdx}
                  className="border border-slate-300 rounded-lg p-3 mb-2 bg-slate-50"
                >
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-slate-700">
                      Shed {sIdx + 1}
                      {shed.shed_name ? ` — ${shed.shed_name}` : ""}
                    </span>
                    <button
                      type="button"
                      onClick={() => removeShed(sIdx)}
                      disabled={formSheds.length === 1}
                      className="text-red-600 hover:text-red-800 disabled:opacity-40 text-[11px] font-semibold"
                    >
                      Remove Shed
                    </button>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mb-2">
                    <div>
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">
                        Shed Name
                      </label>
                      <input
                        placeholder="Shed 1 or Bridge par"
                        value={shed.shed_name || ""}
                        onChange={(e) =>
                          updateShedField(sIdx, "shed_name", e.target.value)
                        }
                        className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                      />
                      
                    </div>
                    <div>
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">
                        Farm Category *
                      </label>
                      <SearchableSelect
                        searchable={false}
                        value={shed.farm_category || ""}
                        onChange={(v) => updateShedField(sIdx, "farm_category", v)}
                        placeholder="Select Category"
                        options={[
                          { value: "Poultry", label: "🐔 Poultry" },
                          { value: "Cattle", label: "🐄 Cattle" },
                          { value: "Fish", label: "🐟 Fish" },
                        ]}
                      />
                    </div>
                  </div>

                  <div className="mb-2 flex flex-wrap items-center gap-2 bg-white border border-slate-200 rounded p-2.5">
                    <button
                      type="button"
                      onClick={() => {
                        setPickingShedIndex(sIdx);
                        setPickerCenterOverride(null);
                        setShowMap(true);
                      }}
                      className="inline-flex items-center gap-1.5 bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-semibold"
                    >
                      <MapPin size={13} /> Select Location from Map
                    </button>
                    <button
                      type="button"
                      onClick={() => handleLocateMeForShed(sIdx)}
                      disabled={locatingInPicker}
                      className="inline-flex items-center gap-1.5 bg-slate-600 text-white px-3 py-1.5 rounded text-xs font-semibold disabled:opacity-50"
                    >
                      <Navigation size={13} /> Locate Me
                    </button>
                    {shed.latitude && shed.longitude && (
                      <span className="text-emerald-700 text-xs font-semibold">
                        📍 Selected ({Number(shed.latitude).toFixed(4)},{" "}
                        {Number(shed.longitude).toFixed(4)})
                      </span>
                    )}
                  </div>

                  {shed.farm_category === "Poultry" && (
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 mb-2 bg-amber-50 border border-amber-200 rounded p-2.5">
                      {POULTRY_FIELDS.map((item) => (
                        <div key={item}>
                          <label className="block text-[11px] font-semibold capitalize">
                            {item.replace("_", " ")}
                          </label>
                          <input
                            type="number"
                            min="0"
                            value={shed[item] || ""}
                            onChange={(e) =>
                              updateShedField(sIdx, item, e.target.value)
                            }
                            className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                          />
                        </div>
                      ))}
                    </div>
                  )}

                  {shed.farm_category === "Cattle" && (
                    <div className="grid grid-cols-3 gap-2 mb-2 bg-emerald-50 border border-emerald-200 rounded p-2.5">
                      {CATTLE_FIELDS.map((item) => (
                        <div key={item}>
                          <label className="block text-[11px] font-semibold capitalize">
                            {item.replace("_", " ")}
                          </label>
                          <input
                            type="number"
                            min="0"
                            value={shed[item] || ""}
                            onChange={(e) =>
                              updateShedField(sIdx, item, e.target.value)
                            }
                            className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                          />
                        </div>
                      ))}
                    </div>
                  )}

                  {shed.farm_category === "Fish" && (
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-2 bg-sky-50 border border-sky-200 rounded p-2.5">
                      {FISH_FIELDS.map((item) => (
                        <div key={item}>
                          <label className="block text-[11px] font-semibold capitalize">
                            {item}
                          </label>
                          <input
                            type="number"
                            min="0"
                            value={shed[item] || ""}
                            onChange={(e) =>
                              updateShedField(sIdx, item, e.target.value)
                            }
                            className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                          />
                        </div>
                      ))}
                    </div>
                  )}

                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 mb-1">
                      Feed Type
                    </label>
                    <input
                      placeholder="Feed Type"
                      value={shed.feed_type || ""}
                      onChange={(e) =>
                        updateShedField(sIdx, "feed_type", e.target.value)
                      }
                      className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                    />
                  </div>
                </div>
              ))}

              <button
                type="button"
                onClick={addShed}
                className="inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold"
              >
                <Plus size={13} /> Add Shed
              </button>
            </div>

            <div className="grid grid-cols-1 gap-3 mb-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Remarks
                </label>
                <input
                  placeholder="Remarks"
                  value={form.remarks}
                  onChange={(e) =>
                    setForm({ ...form, remarks: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t">
              <button
                onClick={() => setShowForm(false)}
                className="px-4 py-1.5 rounded border border-slate-300 text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleSave}
                className="px-4 py-1.5 rounded bg-green-600 text-white text-xs font-semibold"
              >
                {editingId ? "Update Farmer" : "Save Farmer"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 7. LOCATION PICKER MODAL */}
      {showMap && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[2100] p-2 sm:p-4">
          <div className="bg-white p-3 sm:p-4 rounded-xl w-full max-w-4xl shadow-2xl space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="font-bold text-sm sm:text-base">
                Select Shed Location
                {pickingShedIndex != null
                  ? ` — Shed ${pickingShedIndex + 1}`
                  : ""}
              </h2>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleLocateMeInPicker}
                  disabled={locatingInPicker}
                  className="inline-flex items-center gap-1 bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                >
                  <Navigation size={13} />
                  {locatingInPicker ? "Locating..." : "Locate Me"}
                </button>

                {pickingShedIndex != null &&
                  formSheds[pickingShedIndex]?.latitude &&
                  formSheds[pickingShedIndex]?.longitude && (
                    <button
                      type="button"
                      onClick={() => setShowMap(false)}
                      className="inline-flex items-center gap-1 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                    >
                      <CheckCircle2 size={13} /> Use Selected Location
                    </button>
                  )}

                <button onClick={() => setShowMap(false)}>
                  <X size={18} />
                </button>
              </div>
            </div>

            <LocationPicker
              key={
                pickerCenterOverride
                  ? `${pickerCenterOverride[0]}_${pickerCenterOverride[1]}`
                  : "default_picker"
              }
              center={
                pickerCenterOverride
                  ? pickerCenterOverride
                  : formAgentCenter
                  ? formAgentCenter
                  : selectedUpazila
                  ? [Number(selectedUpazila.lat), Number(selectedUpazila.lon)]
                  : selectedDistrict
                  ? [
                      Number(selectedDistrict.lat),
                      Number(selectedDistrict.lon),
                    ]
                  : [24.7471, 90.4203]
              }
              zoom={
                pickerCenterOverride
                  ? 17
                  : formAgentCenter
                  ? 14
                  : selectedUpazila
                  ? 13
                  : selectedDistrict
                  ? 11
                  : 7
              }
              selectedPosition={
                pickingShedIndex != null &&
                formSheds[pickingShedIndex]?.latitude &&
                formSheds[pickingShedIndex]?.longitude
                  ? [
                      Number(formSheds[pickingShedIndex].latitude),
                      Number(formSheds[pickingShedIndex].longitude),
                    ]
                  : null
              }
              onSelect={(lat, lng) => {
                const targetIndex = pickingShedIndex ?? 0;
                setFormSheds((prev) =>
                  prev.map((s, i) =>
                    i === targetIndex
                      ? { ...s, latitude: lat.toString(), longitude: lng.toString() }
                      : s
                  )
                );
                setShowMap(false);
              }}
              markers={
                <>
                  {farmers.flatMap((farmer: any) => {
                    const sheds = getShedsForFarmer(farmer);
                    return sheds
                      .map((shed: any, sIdx: number) => ({
                        farmer,
                        shed,
                        sIdx,
                        totalSheds: sheds.length,
                      }))
                      .filter((x: any) => x.shed.latitude && x.shed.longitude)
                      .map(({ farmer: f2, shed, sIdx, totalSheds }: any) => (
                        <Marker
                          key={
                            shed.id ? `shed-${shed.id}` : `${f2.id}-${sIdx}`
                          }
                          position={[
                            Number(shed.latitude),
                            Number(shed.longitude),
                          ]}
                          icon={buildCustomFarmMarker(
                            f2,
                            shed,
                            false,
                            undefined,
                            sIdx,
                            totalSheds
                          )}
                        />
                      ));
                  })}
                </>
              }
            />
          </div>
        </div>
      )}

      {/* 8. FARMER SUMMARY MODAL */}
      {showSummary && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-6 w-full max-w-6xl max-h-[90vh] overflow-auto">
            <div className="flex justify-between items-center pb-3 mb-4 border-b">
              <div>
                <h2 className="text-base sm:text-lg font-bold">
                  Farmer Summary — {selectedCompanyName}
                </h2>
                <p className="text-xs text-slate-500">
                  District: {districtFilter || "All"} | Upazila:{" "}
                  {upazilaFilter || "All"} | Union: {unionFilter || "All"}
                </p>
              </div>
              <button onClick={() => setShowSummary(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="mb-4">
              <select
                defaultValue=""
                onChange={(e) => {
                  const val = e.target.value;
                  if (val === "company") exportFarmerSummaryExcel();
                  if (val === "poultry") exportPoultrySummaryExcel();
                  if (val === "cattle") exportCattleSummaryExcel();
                  if (val === "fish") exportFishSummaryExcel();
                  e.target.value = "";
                }}
                className="bg-slate-800 text-white px-3.5 py-1.5 rounded text-xs font-semibold cursor-pointer"
              >
                <option value="" hidden>
                  📊 Export Summary
                </option>
                <option value="company">Company Farmer Summary</option>
                <option value="poultry">🐔 Poultry Summary</option>
                <option value="cattle">🐄 Cattle Summary</option>
                <option value="fish">🐟 Fish Summary</option>
              </select>
            </div>

            <div className="overflow-x-auto mb-6">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-slate-100">
                  <tr>
                    <th className="border px-3 py-1.5 text-left">Company</th>
                    <th className="border px-3 py-1.5 text-center">
                      Total Farmers
                    </th>
                    <th className="border px-3 py-1.5 text-center">
                      🐔 Poultry Farmers
                    </th>
                    <th className="border px-3 py-1.5 text-center">
                      🐄 Cattle Farmers
                    </th>
                    <th className="border px-3 py-1.5 text-center">
                      🐟 Fish Farmers
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((row: any) => (
                    <tr key={row.company}>
                      <td className="border px-3 py-1.5 font-semibold">
                        {row.company}
                      </td>
                      <td className="border px-3 py-1.5 text-center font-bold">
                        {row.totalFarmers}
                      </td>
                      <td className="border px-3 py-1.5 text-center text-amber-700 font-semibold">
                        {row.poultryFarmers}
                      </td>
                      <td className="border px-3 py-1.5 text-center text-emerald-700 font-semibold">
                        {row.cattleFarmers}
                      </td>
                      <td className="border px-3 py-1.5 text-center text-sky-700 font-semibold">
                        {row.fishFarmers}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 className="text-sm font-bold mb-2">🐔 Poultry Summary</h3>
            <div className="overflow-x-auto mb-6">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-amber-100">
                  <tr>
                    <th className="border px-2 py-1.5 text-left">Company</th>
                    <th className="border px-2 py-1.5">Broiler Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Broiler
                    </th>
                    <th className="border px-2 py-1.5">Layer Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Layer
                    </th>
                    <th className="border px-2 py-1.5">Sonali Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Sonali
                    </th>
                    <th className="border px-2 py-1.5">Duck Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Duck
                    </th>
                    <th className="border px-2 py-1.5">Quail Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Quail
                    </th>
                    <th className="border px-2 py-1.5">Deshi Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Deshi
                    </th>
                    <th className="border px-2 py-1.5 bg-amber-200 font-bold">
                      Total Poultry
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((row: any) => (
                    <tr key={row.company}>
                      <td className="border px-2 py-1 font-semibold">
                        {row.company}
                      </td>
                      <td className="border text-center">
                        {row.broilerFarmers}
                      </td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.broiler}
                      </td>
                      <td className="border text-center">{row.layerFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.layer}
                      </td>
                      <td className="border text-center">
                        {row.sonaliFarmers}
                      </td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.sonali}
                      </td>
                      <td className="border text-center">{row.duckFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.duck}
                      </td>
                      <td className="border text-center">{row.quailFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.quail}
                      </td>
                      <td className="border text-center">{row.deshiFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.deshiChicken}
                      </td>
                      <td className="border text-center bg-amber-100 font-bold">
                        {row.totalPoultry}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 className="text-sm font-bold mb-2">🐄 Cattle Summary</h3>
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-emerald-100">
                  <tr>
                    <th className="border px-2 py-1.5 text-left">Company</th>
                    <th className="border px-2 py-1.5">Dairy Farms</th>
                    <th className="border px-2 py-1.5 bg-emerald-50">
                      Total Dairy Cow
                    </th>
                    <th className="border px-2 py-1.5">Bull Farms</th>
                    <th className="border px-2 py-1.5 bg-emerald-50">
                      Total Bull
                    </th>
                    <th className="border px-2 py-1.5">Calf/Heifer Farms</th>
                    <th className="border px-2 py-1.5 bg-emerald-50">
                      Total Calf/Heifer
                    </th>
                    <th className="border px-2 py-1.5 bg-emerald-200 font-bold">
                      Total Cattle
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((row: any) => (
                    <tr key={row.company}>
                      <td className="border px-2 py-1 font-semibold">
                        {row.company}
                      </td>
                      <td className="border text-center">
                        {row.dairyCowFarmers}
                      </td>
                      <td className="border text-center bg-emerald-50 font-semibold">
                        {row.dairyCow}
                      </td>
                      <td className="border text-center">{row.bullFarmers}</td>
                      <td className="border text-center bg-emerald-50 font-semibold">
                        {row.bull}
                      </td>
                      <td className="border text-center">
                        {row.calfHeiferFarmers}
                      </td>
                      <td className="border text-center bg-emerald-50 font-semibold">
                        {row.calfHeifer}
                      </td>
                      <td className="border text-center bg-emerald-100 font-bold">
                        {row.totalCattle}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}