// ================= PART 1 OF 2 =================
import React, { useEffect, useRef, useState } from "react";
import { useLocation } from "react-router-dom";
import { supabase } from "../../lib/supabase";
import { useAuth } from "../../context/AuthContext";
import {
  canEditRecord,
  canAddReport,
  canAddFarmer,
} from "../../utils/accessControl";
import LocationPicker from "../../components/maps/LocationPicker";
import { Marker } from "react-leaflet";
import L from "leaflet";
import {
  MapPin,
  Navigation,
  CheckCircle2,
  X,
  Plus,
  UserPlus,
  Trash2,
} from "lucide-react";
import ExcelJS from "exceljs";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Capacitor } from "@capacitor/core";
import SearchableSelect from "../../components/common/SearchableSelect";


export const CATEGORIES = [
  "Broiler",
  "Layer",
  "Sonali",
  "Duck",
  "Quail",
  "Cattle",
  "Fish",
];

export const BREEDS_BY_CATEGORY: Record<string, string[]> = {
  Broiler: [
    "Arber Acress",
    "Ross",
    "Cobb 500",
    "Indian River",
    "Hubbard",
    "Hubbard EP",
  ],
  Layer: [
    "Shaver Star Cross 579 Brown",
    "Bovans White",
    "Bovans Brown",
    "Lohman Brown",
    "Lohman White",
    "Isa Brown",
    "Hisex Brown",
    "Hisex White",
    "Novojen Brown",
    "Hyline Brown",
    "Hyline White",
    "Decalb White",
    "BB 300",
  ],
  Sonali: [
    "Local",
    "Classic",
    "Hybrid",
    "Dominant",
    "Hubbard Premium",
    "Sasso",
    "Color Bird/Sasso",
    "Tetra HB",
    "Bornali",
    "Bornali Plus",
    "Cock",
    "Brown Cock",
    "White Cock",
  ],
  Duck: ["Khaki Campbell", "Pekin", "Jinding", "Deshi Duck"],
  Quail: ["Japanese Quail", "Bobwhite Quail"],
  Cattle: ["Holstein Friesian", "Sahiwal", "Sindhi", "Deshi", "Crossbred"],
  Fish: ["Pangus", "Tilapia", "Koi", "Pabda", "Gulsha", "Carp", "Mixed"],
};

export const FEED_TYPES_BY_CATEGORY: Record<string, string[]> = {
  Broiler: ["Broiler Pre Starter", "Starter", "Grower", "Finisher"],
  Layer: [
    "Broiler Pre Starter",
    "Starter",
    "Sonali Starter",
    "Layer Starter",
    "Grower",
    "Layer Layer 1",
    "Layer Layer 2",
    "Corporate Layer",
    "Layer Concentrated",
  ],
  Sonali: [
    "Sonali Starter",
    "Grower",
    "Finisher",
    "Broiler Pre Starter",
    "Starter",
  ],
};

type StatusTab =
  | "all"
  | "Running"
  | "Active"
  | "Inactive"
  | "Leave from Nourish"
  | "Convert to Nourish";

const STOCK_KEY_BY_CATEGORY: Record<string, string> = {
  Broiler: "broiler",
  Layer: "layer",
  Sonali: "sonali",
  Duck: "duck",
  Quail: "quail",
};
const POULTRY_STOCK_KEYS = ["broiler", "layer", "sonali", "duck", "quail"];
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

const EMPTY_FARMER_FORM = {
  company_id: "",
  agent_id: "",
  // How the farmer is assigned: "agent" = directly under an agent,
  // "sub" = under a sub-agent (agent_id then holds the sub-agent's parent).
  assign_to: "agent",
  sub_agent_id: "",
  farmer_name: "",
  mobile: "",
  village: "",
  union_name: "",
  upazila: "",
  district: "",
  farm_category: "Poultry",
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

const STATUS_RANK: Record<string, number> = {
  Running: 0,
  Active: 1,
  Inactive: 2,
};

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

function normalizeCategory(f: any): string {
  const raw = String(f.farm_type || f.farm_category || f.category || "").trim();
  const matched = CATEGORIES.find((c) => c.toLowerCase() === raw.toLowerCase());
  if (matched) return matched;
  if (Number(f.layer || 0) > 0) return "Layer";
  if (Number(f.sonali || 0) > 0) return "Sonali";
  if (Number(f.duck || 0) > 0) return "Duck";
  if (Number(f.quail || 0) > 0) return "Quail";
  if (f.farm_category === "Cattle") return "Cattle";
  if (f.farm_category === "Fish") return "Fish";
  return "Broiler";
}

function getDaysSince(dateStr?: string | null): number | null {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return null;
  return Math.floor((new Date().getTime() - d.getTime()) / 86400000);
}

function batchTagOf(label?: string, farmerName?: string): string {
  if (!label) return "";
  if (label === farmerName) return "";
  if (farmerName && label.startsWith(farmerName)) {
    const rest = label.slice(farmerName.length).replace(/^[\s\-–:]+/, "").trim();
    return rest || label;
  }
  return label;
}


const SHED_STOCK_FIELDS = [
  ...POULTRY_FIELDS,
  ...CATTLE_FIELDS,
  ...FISH_FIELDS,
];

const EMPTY_SHED_FORM: any = {
  id: undefined,
  shed_name: "",
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
};

// ---------- Multi-shed helpers ----------
// A farmer's sheds come from farmer_sheds when that table exists; otherwise
// one legacy shed is synthesised from the farmer row. The first/only shed
// has an EMPTY shed_name — it is shown as just the farmer name.
function getShedsForFarmer(farmer: any): any[] {
  const list = farmer?.farmer_sheds || farmer?.sheds || [];
  if (Array.isArray(list) && list.length > 0) return list;
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
  SHED_STOCK_FIELDS.forEach((k) => {
    shed[k] = Number(farmer?.[k] || 0);
  });
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

// The shed a visit belongs to: its stored shed_id when it matches one of
// the farmer's sheds; the only shed when the farmer has exactly one; null
// when the farmer has several sheds and the visit names none (do not guess).
function getShedForVisit(visit: any, farmer: any): any | null {
  if (!farmer) return null;
  const sheds = getShedsForFarmer(farmer);
  if (sheds.length === 0) return null;
  const sid = visit?.shed_id;
  if (sid !== null && sid !== undefined && String(sid) !== "") {
    return sheds.find((s: any) => String(s.id) === String(sid)) || null;
  }
  return sheds.length === 1 ? sheds[0] : null;
}

function getVisitShedDisplayName(visit: any, farmer: any): string {
  const farmerName =
    farmer?.farmer_name ||
    visit?.farmers?.farmer_name ||
    visit?.batch_label ||
    "Farmer";
  if (!farmer) return farmerName;
  const sheds = getShedsForFarmer(farmer);
  const shed = getShedForVisit(visit, farmer);
  if (!shed) return farmerName;
  return getShedDisplayName(farmer, shed, sheds.indexOf(shed), sheds.length);
}

// Per-shed status — the same rule as the farmer-level computedStatus,
// but keyed by shed (identical helper in FarmersPage): a shed with a
// Running/Activated visit is Running; otherwise the status comes from
// that shed's own last sale (<=60 days => Active, older => Inactive);
// a shed with no visit history falls back to its stored status
// (default Active). A visit without a shed_id counts only when the
// farmer has a single shed (same rule as getShedForVisit).
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
    // No visit history at all: a stored "Running" is stale (it survives
    // when the running batch's visit row is deleted), so only a stored
    // Active/Inactive is trusted — otherwise the shed is Active.
    status =
      shed?.status === "Active" || shed?.status === "Inactive"
        ? shed.status
        : "Active";
  return {
    status,
    statusDate: lastSale || "",
    daysSinceSale: days,
    closedType: closed[0]?.status || "",
  };
}

function shedFormToDb(
  shed: any,
  farmerId: number,
  farmerName: string = "",
  index: number = 0,
  totalSheds: number = 1
): any {
  const db: any = {
    farmer_id: farmerId,
    shed_name: getShedShortName(
      { farmer_name: farmerName },
      shed,
      index,
      totalSheds
    ),
    farm_category: shed?.farm_category || "",
    latitude:
      shed?.latitude === "" || shed?.latitude == null
        ? null
        : Number(shed.latitude),
    longitude:
      shed?.longitude === "" || shed?.longitude == null
        ? null
        : Number(shed.longitude),
    feed_type: shed?.feed_type || "",
    status: shed?.status || "Active",
  };
  SHED_STOCK_FIELDS.forEach((k) => {
    db[k] = Number(shed?.[k] || 0);
  });
  return db;
}

function hasValue(val: any): boolean {
  if (val === null || val === undefined) return false;
  const s = String(val).trim();
  if (s === "" || s === "-" || s === "null" || s === "undefined") return false;
  return true;
}

function getRemarkText(rep: any): string {
  const r = rep?.remarks ?? rep?.remark ?? "";
  return String(r || "").trim();
}

export const UPDATED_FEED_TYPES_BY_CATEGORY: Record<string, string[]> = {
  Broiler: [
    "Broiler Pre Starter",
    "Broiler Starter",
    "Broiler Grower",
    "Broiler Finisher",
  ],
  Layer: [
    "Broiler Pre Starter",
    "Broiler Starter",
    "Sonali Starter",
    "Layer Starter",
    "Layer Grower",
    "Layer Layer 1",
    "Layer Layer 2",
    "Corporate Layer",
    "Layer Concentrated",
  ],
  Sonali: [
    "Sonali Starter",
    "Sonali Grower",
    "Sonali Finisher",
    "Broiler Pre Starter",
    "Broiler Starter",
    "Broiler Grower",
    "Broiler Finisher",
  ],
};

const DEFAULT_VACCINES = [
  {
    category: "Broiler",
    age_days: 5,
    vaccine_name: "IB + ND",
    vaccine_type: "Live",
    route: "Eye Drop",
  },
  {
    category: "Broiler",
    age_days: 11,
    vaccine_name: "IBD Intermediate Strain (11-12d)",
    vaccine_type: "Live",
    route: "Eye Drop",
  },
  {
    category: "Broiler",
    age_days: 18,
    vaccine_name: "IBD Intermediate Plus Strain (18-19d)",
    vaccine_type: "Live",
    route: "Eye Drop/Drinking Water",
  },
  {
    category: "Broiler",
    age_days: 22,
    vaccine_name: "ND Lasota (22-23d)",
    vaccine_type: "Live",
    route: "Drinking Water",
  },
  {
    category: "Sonali",
    age_days: 5,
    vaccine_name: "IB + ND",
    vaccine_type: "Live",
    route: "Eye Drop",
  },
  {
    category: "Sonali",
    age_days: 11,
    vaccine_name: "IBD Intermediate Strain (11-12d)",
    vaccine_type: "Live",
    route: "Eye Drop",
  },
  {
    category: "Sonali",
    age_days: 18,
    vaccine_name: "IBD Intermediate Plus Strain (18-19d)",
    vaccine_type: "Live",
    route: "Eye Drop/Drinking Water",
  },
  {
    category: "Sonali",
    age_days: 22,
    vaccine_name: "ND Lasota (22-23d)",
    vaccine_type: "Live",
    route: "Drinking Water",
  },
  {
    category: "Layer",
    age_days: 5,
    vaccine_name: "IB + ND",
    vaccine_type: "Live",
    route: "Eye Drop",
  },
  {
    category: "Layer",
    age_days: 12,
    vaccine_name: "Gumboro (IBD) 1st Dose",
    vaccine_type: "Live",
    route: "Eye Drop",
  },
  {
    category: "Layer",
    age_days: 20,
    vaccine_name: "Gumboro (IBD) 2nd Dose",
    vaccine_type: "Live",
    route: "Drinking Water",
  },
  {
    category: "Layer",
    age_days: 28,
    vaccine_name: "ND Lasota",
    vaccine_type: "Live",
    route: "Eye Drop / Drinking Water",
  },
  {
    category: "Layer",
    age_days: 35,
    vaccine_name: "Fowl Pox",
    vaccine_type: "Live",
    route: "Wing Web",
  },
  {
    category: "Layer",
    age_days: 56,
    vaccine_name: "ND Killed / RDV",
    vaccine_type: "Killed",
    route: "IM Injection",
  },
];

function formatDmy(dateStr?: string) {
  if (!dateStr) return "-";
  const parts = String(dateStr).slice(0, 10).split("-");
  if (parts.length !== 3) return dateStr;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function isWhiteLayerBreedName(breedName?: string): boolean {
  const b = String(breedName || "").trim().toLowerCase();
  return (
    b.includes("white") ||
    b === "bo" ||
    b.includes("bovans white") ||
    b.includes("lohman white") ||
    b.includes("hisex white") ||
    b.includes("hyline white") ||
    b.includes("decalb") ||
    b.includes("bb 300")
  );
}

export default function VetVisitsPage() {
  const { role } = useAuth();
  const isSuperAdmin = role === "super_admin";
  const isFullAccess = role === "super_admin" || role === "manager";
  // Permissions (src/utils/accessControl.ts): this page is the POULTRY
  // visit page (rule 5) — report add/edit/delete needs Poultry
  // assigned; edits/deletes also need ownership (old rows with
  // created_by null are editable by area officers — the save stamps
  // the creator). Farmer add is NOT category-gated. View / history /
  // export / Map stay open to everyone.
  const [currentUserId, setCurrentUserId] = useState("");
  const [assignedCategories, setAssignedCategories] = useState<string[]>([]);
  const canAddPoultryReport = canAddReport(role, assignedCategories, "Poultry");
  const canAddFarmerHere = canAddFarmer(role);
  const canEditPoultryReport = (rec: any) =>
    canAddPoultryReport && canEditRecord(role, rec?.created_by, currentUserId);
  // Full row edit (farmer + batch modal): farmer part follows the
  // farmer rule, batch part follows the Poultry report rule.
  const canFullEditRow = (r: any) =>
    canEditRecord(role, r?.farmer?.created_by, currentUserId) &&
    (!r?.visit || canEditPoultryReport(r.visit));

  const [currentUserProfile, setCurrentUserProfile] = useState({
    fullName: "Officer",
    designation: "Customer Service Officer",
    region: "Mymensingh Zone",
  });
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);
  const [assignedDistricts, setAssignedDistricts] = useState<string[]>([]);

  const [companies, setCompanies] = useState<any[]>([]);
  const [allAgents, setAllAgents] = useState<any[]>([]);
  const [subAgents, setSubAgents] = useState<any[]>([]);
  const [selectedCompany, setSelectedCompany] = useState<string>("all");
  const [selectedUpazila, setSelectedUpazila] = useState<string>("all");
  const [selectedAgent, setSelectedAgent] = useState<string>("all");
  const [selectedSubAgent, setSelectedSubAgent] = useState<string>("all");
  const [selectedCategoryFilter, setSelectedCategoryFilter] =
    useState<string>("all");
  const [selectedStatusFilter, setSelectedStatusFilter] =
    useState<StatusTab>("all");
  const [searchQuery, setSearchQuery] = useState<string>("");

  const [_visitsData, setVisitsData] = useState<any[]>([]);
  const location = useLocation();
  const companyInitRef = useRef(false);
  const [statusFarmers, setStatusFarmers] = useState<any[]>([]);
  const [expandedCard, setExpandedCard] = useState<string | null>(null);
  const [focusVisitId, setFocusVisitId] = useState<string | null>(null);
  const [focusFarmerId, setFocusFarmerId] = useState<string | null>(null);

  const [visitReportMeta, setVisitReportMeta] = useState<
    Record<string, { count: number; lastDate: string }>
  >({});
  const [farmerReportMeta, setFarmerReportMeta] = useState<
    Record<string, { count: number; lastDate: string }>
  >({});

  // Activate / Run Farm / Add Batch
  const [showActivateModal, setShowActivateModal] = useState(false);
  const [activateFarmer, setActivateFarmer] = useState<any>(null);
  const [activateMode, setActivateMode] = useState<"add" | "replace">("add");
  const [activateForm, setActivateForm] = useState({
    category: "Broiler",
    hatch_date: new Date().toISOString().split("T")[0],
    hatchery: "",
    breed: "",
    grade: "AG",
    box_no: "",
    t_birds: "",
    batch_mark: "",
  });

  // Add / Full Edit Farmer + Batch Modal
  const [showFarmerFormModal, setShowFarmerFormModal] = useState(false);
  const [editingFarmerId, setEditingFarmerId] = useState<number | null>(null);
  const [editingBatchVisit, setEditingBatchVisit] = useState<any>(null);
  const [farmerForm, setFarmerForm] = useState({ ...EMPTY_FARMER_FORM });
  const [batchEditSubForm, setBatchEditSubForm] = useState({
    hatch_date: "",
    hatchery: "",
    breed: "",
    grade: "AG",
    box_no: "",
    t_birds: "",
    category: "Broiler",
    batch_label: "",
  });
  // Multi-shed: chosen shed for the Activate/Run Farm modal, and the shed
  // list edited in the Add/Edit Farmer modal (first shed = farmer form).
  const [activateShedId, setActivateShedId] = useState<string>("");
  const [addingActivateShed, setAddingActivateShed] = useState(false);
  const [savingFarmer, setSavingFarmer] = useState(false);
  const [savingActivate, setSavingActivate] = useState(false);
  const [savingVaccine, setSavingVaccine] = useState(false);
  const [savingBefore, setSavingBefore] = useState(false);
  const [savingAfter, setSavingAfter] = useState(false);
  const [savingLayer, setSavingLayer] = useState(false);
  const [savingAddVaccine, setSavingAddVaccine] = useState(false);
  const [formSheds, setFormSheds] = useState<any[]>([{ ...EMPTY_SHED_FORM }]);
  const [geoDistricts, setGeoDistricts] = useState<any[]>([]);
  const [geoUpazilas, setGeoUpazilas] = useState<any[]>([]);
  const [geoUnions, setGeoUnions] = useState<any[]>([]);
  const [showMapPicker, setShowMapPicker] = useState(false);
  const [pickerCenterOverride, setPickerCenterOverride] = useState<
    [number, number] | null
  >(null);
  const [locatingInPicker, setLocatingInPicker] = useState(false);

  // Track if a report modal is currently editing an existing report from History
  const [editingReportId, setEditingReportId] = useState<any>(null);
  const [editingReportCreatedBy, setEditingReportCreatedBy] = useState<string | null>(null);

  const [vaccines, setVaccines] = useState<any[]>([]);
  const [broilerStds, setBroilerStds] = useState<Record<number, any>>({});
  const [sonaliStds, setSonaliStds] = useState<Record<number, any>>({});
  const [layerStds, setLayerStds] = useState<Record<string, any>>({});
  const [_loading, setLoading] = useState(true);

  // Master Vaccine Schedule Modal
  const [showVaccineModal, setShowVaccineModal] = useState(false);
  const [vaccineForm, setVaccineForm] = useState({
    category: "Broiler",
    age_days: "",
    vaccine_name: "",
    vaccine_type: "Live",
    route: "Eye Drop",
  });

  // Quick "+" Log Given Vaccine Modal
  const [showLogVaccineModal, setShowLogVaccineModal] = useState(false);
  const [activeVaccineVisit, setActiveVaccineVisit] = useState<any>(null);
  const [logVaccineForm, setLogVaccineForm] = useState({
    vaccine_name: "",
    vaccine_date: new Date().toISOString().split("T")[0],
  });

  // History Modal (Selected batch upfront + other batches in dropdown)
  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [selectedHistoryVisit, setSelectedHistoryVisit] = useState<any>(null);
  const [historyBatches, setHistoryBatches] = useState<any[]>([]);
  const [allReportsForFarmer, setAllReportsForFarmer] = useState<{
    before: any[];
    after: any[];
    layer: any[];
    feedLogs: any[];
  }>({ before: [], after: [], layer: [], feedLogs: [] });

  // Report Modals
  const [showBeforeSaleModal, setShowBeforeSaleModal] = useState(false);
  const [showAfterSaleModal, setShowAfterSaleModal] = useState(false);
  const [showLayerModal, setShowLayerModal] = useState(false);
  const [activeReportVisit, setActiveReportVisit] = useState<any>(null);
  const [reportVisitDate, setReportVisitDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );

  const [existingReportOnDate, setExistingReportOnDate] = useState(false);
  const [checkingDate, setCheckingDate] = useState(false);

  useEffect(() => {
    async function checkDate() {
      if (editingReportId || !activeReportVisit || !reportVisitDate) {
        setExistingReportOnDate(false);
        return;
      }
      let table = "";
      if (showBeforeSaleModal) table = "before_sale_reports";
      else if (showAfterSaleModal) table = "after_sale_reports";
      else if (showLayerModal) table = "layer_reports";

      if (!table) {
        setExistingReportOnDate(false);
        return;
      }

      setCheckingDate(true);
      const { data } = await supabase
        .from(table)
        .select("id")
        .eq("vet_visit_id", activeReportVisit.id)
        .eq("visit_date", reportVisitDate)
        .limit(1)
        .maybeSingle();

      setExistingReportOnDate(!!data);
      setCheckingDate(false);
    }
    checkDate();
  }, [
    activeReportVisit,
    reportVisitDate,
    showBeforeSaleModal,
    showAfterSaleModal,
    showLayerModal,
    editingReportId,
  ]);


  const [transferAgentId, setTransferAgentId] = useState<string>("");

  // Layer Culled Reason States
  const [showCullReasonInput, setShowCullReasonInput] = useState(false);
  const [cullReasonText, setCullReasonText] = useState("");

  // WhatsApp Popup State
  const [whatsappPopupText, setWhatsappPopupText] = useState<string>("");
  const [copiedWa, setCopiedWa] = useState(false);

  // Before Sale Form
  const [beforeForm, setBeforeForm] = useState({
    mortality_pcs: "",
    weight_gm_b: "",
    feed_t_kg: "",
    feed_gm_b: "",
    fcr_w_out_m: "",
    fcr_with_m: "",
    uneven: "",
    indigestion: "",
    orange_feces: "",
    watery_feces: "",
    feeder: "",
    drinker: "",
    floor_space: "",
    litter_condition: "",
    ventilation: "",
    feed_scattering: "",
    feed_hardness: "",
    feed_company: "",
    feed_type: "",
    pro_date: "",
    batch_no: "",
    remarks: "",
  });

  // After Sale Form
  const [afterForm, setAfterForm] = useState({
    mortality: "",
    weight: "",
    feed_consumption_kg: "",
    feed_gm_b: "",
    feed: "",
    feed_mill: "",
    fcr_w_out_m: "",
    fcr_with_m: "",
    remarks: "",
  });

  // Layer Form
  const [layerForm, setLayerForm] = useState({
    mortality_pcs: "0",
    b_weight_g: "",
    feed_total_kg: "",
    feed_intake_g: "",
    egg_production: "",
    egg_weight: "",
    feed_company: "",
    feed_type: "",
    pro_date: "",
    batch_no: "",
    hatchery: "",
    breed: "",
    disease: "",
    remarks: "",
  });

  useEffect(() => {
    loadInitialData();
  }, []);

  useEffect(() => {
    if (companyInitRef.current || companies.length === 0) return;
    companyInitRef.current = true;
    const focusParams = new URLSearchParams(location.search);
    const hasFocus = focusParams.get("focus") || focusParams.get("focusFarmer");
    if (hasFocus) return;
    const nourish = companies.find((c) =>
      c.company_name?.toLowerCase().includes("nourish")
    );
    if (nourish) setSelectedCompany(String(nourish.id));
  }, [companies, location.search]);

  useEffect(() => {
    const focusParams = new URLSearchParams(location.search);
    const fid = focusParams.get("focus");
    const farmerFid = focusParams.get("focusFarmer");
    if (!fid && !farmerFid) return;
    if (fid) setFocusVisitId(fid);
    if (farmerFid) setFocusFarmerId(farmerFid);
    setSelectedCompany("all");
    setSelectedUpazila("all");
    setSelectedAgent("all");
    setSelectedStatusFilter("all");
    setSelectedCategoryFilter("all");
    setSearchQuery("");
  }, [location.search]);

  useEffect(() => {
    if ((!focusVisitId && !focusFarmerId) || _loading) return;
    const t = setTimeout(() => {
      const el =
        (focusVisitId &&
          (document.getElementById("after-sale-btn-" + focusVisitId) ||
            document.getElementById("visit-row-" + focusVisitId))) ||
        (focusFarmerId &&
          document.querySelector(`[data-farmer-id="${focusFarmerId}"]`)) ||
        null;
      el?.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 400);
    return () => clearTimeout(t);
  }, [focusVisitId, focusFarmerId, statusFarmers, _loading]);

  const selectedFormDistrict = geoDistricts.find(
    (d) => d.name === farmerForm.district
  );

  useEffect(() => {
    async function loadModalUpazilas() {
      if (!selectedFormDistrict) {
        setGeoUpazilas([]);
        return;
      }
      const { data } = await supabase
        .from("upazilas")
        .select("*")
        .eq("district_id", selectedFormDistrict.id)
        .order("name");
      const list = data || [];
      const allowed = isFullAccess
        ? list
        : list.filter(
            (u: any) =>
              assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
          );
      setGeoUpazilas(allowed);
    }
    loadModalUpazilas();
  }, [selectedFormDistrict, isFullAccess, assignedUpazilas]);

  const selectedFormUpazila = geoUpazilas.find(
    (u) => u.name === farmerForm.upazila
  );

  // Map picker centre: in Sub Agent mode centre on the sub-agent
  // (farmers cluster around their sub-agent); otherwise the parent
  // agent; falls back to upazila/district when neither has lat/lon.
  const subAgentById: Record<string, any> = {};
  subAgents.forEach((s: any) => {
    subAgentById[String(s.id)] = s;
  });
  const formAgent = farmerForm.agent_id
    ? allAgents.find((a: any) => String(a.id) === String(farmerForm.agent_id))
    : null;
  const formSubAgent =
    farmerForm.assign_to === "sub" && farmerForm.sub_agent_id
      ? subAgentById[String(farmerForm.sub_agent_id)]
      : null;
  const formAgentCenter: [number, number] | null =
    formSubAgent?.latitude && formSubAgent?.longitude
      ? [Number(formSubAgent.latitude), Number(formSubAgent.longitude)]
      : formAgent?.latitude && formAgent?.longitude
      ? [Number(formAgent.latitude), Number(formAgent.longitude)]
      : null;

  // Sub-agents: a farmer always keeps agent_id = the parent agent;
  // sub_agent_id is optional (empty = directly under the agent).
  // Sub-agent option label: "SubAgentName (ParentAgentName)" so the
  // parent agent is visible in the all-sub-agents dropdown.
  function getSubAgentOptionLabel(s: any): string {
    const parent = allAgents.find(
      (a: any) => String(a.id) === String(s.agent_id)
    );
    return parent?.agent_name
      ? `${s.sub_agent_name} (${parent.agent_name})`
      : s.sub_agent_name;
  }

  useEffect(() => {
    async function loadModalUnions() {
      if (!selectedFormUpazila) {
        setGeoUnions([]);
        return;
      }
      const { data } = await supabase
        .from("unions")
        .select("*")
        .eq("upazila_id", selectedFormUpazila.id)
        .order("name");
      setGeoUnions(data || []);
    }
    loadModalUnions();
  }, [selectedFormUpazila]);

  async function loadInitialData() {
    setLoading(true);
    const {
      data: { user },
    } = await supabase.auth.getUser();

    let uUpazilas: string[] = [];
    let uDistricts: string[] = [];
    let uCats: string[] = [];

    if (user) {
      setCurrentUserId(user.id);
      // assigned_categories comes from the permissions migration;
      // fall back gracefully if the column is not there yet.
      let uProf: any = null;
      {
        const { data, error } = await supabase
          .from("users")
          .select("full_name, email, designation, region, upazilas, districts, assigned_categories")
          .eq("id", user.id)
          .single();
        if (!error && data) {
          uProf = data;
        } else {
          const retry = await supabase
            .from("users")
            .select("full_name, email, designation, region, upazilas, districts")
            .eq("id", user.id)
            .single();
          uProf = retry.data || null;
        }
      }

      if (uProf) {
        setCurrentUserProfile({
          fullName: uProf.full_name || uProf.email?.split("@")[0] || "Officer",
          designation:
            uProf.designation || "Customer Service Officer (Poultry)",
          region: uProf.region || "Mymensingh Zone",
        });
        uUpazilas = parseList(uProf.upazilas);
        uDistricts = parseList(uProf.districts);
        uCats = parseList(uProf.assigned_categories);
      }
    }
    setAssignedCategories(uCats);

    setAssignedUpazilas(uUpazilas);
    setAssignedDistricts(uDistricts);
    if (!isFullAccess && uUpazilas.length === 1) {
      setSelectedUpazila(uUpazilas[0]);
    }

    const [
      { data: compData },
      { data: agData },
      { data: vacData },
      { data: bStdData },
      { data: lStdData },
      { data: distData },
    ] = await Promise.all([
      supabase.from("companies").select("*").order("company_name"),
      supabase.from("agents").select("*").order("agent_name"),
      supabase
        .from("vaccines")
        .select("*")
        .order("age_days", { ascending: true }),
      supabase.from("broiler_standards").select("*"),
      supabase.from("layer_standards").select("*"),
      supabase.from("districts").select("*").order("name"),
    ]);

    try {
      const { data: sStdData } = await supabase
        .from("sonali_standards")
        .select("*");
      const sMap: Record<number, any> = {};
      (sStdData || []).forEach((s: any) => {
        if (s.age_days !== undefined) sMap[Number(s.age_days)] = s;
      });
      setSonaliStds(sMap);
    } catch {
      setSonaliStds({});
    }

    const allDists = distData || [];
    const allowedDists = isFullAccess
      ? allDists
      : allDists.filter(
          (d: any) => uDistricts.length === 0 || uDistricts.includes(d.name)
        );
    setGeoDistricts(allowedDists);

    const bMap: Record<number, any> = {};
    (bStdData || []).forEach((s: any) => (bMap[Number(s.age_days)] = s));
    setBroilerStds(bMap);

    // Smart Layer Standards Indexing (Supports breed/type column OR auto-detects White vs Brown by weight when multiple rows exist per week)
    const lMap: Record<string, any> = {};
    const byWeek: Record<number, any[]> = {};

    (lStdData || []).forEach((s: any) => {
      const w = Number(s.age_weeks);
      if (!byWeek[w]) byWeek[w] = [];
      byWeek[w].push(s);

      const rawTag = String(
        s.breed || s.breed_name || s.layer_type || s.type || s.strain || ""
      )
        .trim()
        .toLowerCase();

      if (rawTag) {
        lMap[`${rawTag}_${w}`] = s;
        if (isWhiteLayerBreedName(rawTag)) {
          lMap[`white_${w}`] = s;
        } else if (rawTag.includes("brown") || rawTag.includes("shaver")) {
          lMap[`brown_${w}`] = s;
        }
      }
      if (!lMap[`any_${w}`]) {
        lMap[`any_${w}`] = s;
      }
    });

    // If a week has 2+ rows without explicit white/brown labels, lighter weight = White, heavier weight = Brown
    Object.entries(byWeek).forEach(([wkStr, rows]) => {
      const w = Number(wkStr);
      if (rows.length >= 2 && (!lMap[`white_${w}`] || !lMap[`brown_${w}`])) {
        const sortedByWt = [...rows].sort(
          (a, b) => Number(a.weight || 0) - Number(b.weight || 0)
        );
        if (!lMap[`white_${w}`]) lMap[`white_${w}`] = sortedByWt[0];
        if (!lMap[`brown_${w}`])
          lMap[`brown_${w}`] = sortedByWt[sortedByWt.length - 1];
      }
    });

    setLayerStds(lMap);

    if (!vacData || vacData.length === 0) {
      const { data: insertedVac } = await supabase
        .from("vaccines")
        .insert(DEFAULT_VACCINES)
        .select();
      setVaccines(insertedVac || DEFAULT_VACCINES);
    } else {
      setVaccines(vacData);
    }

    const comps = compData || [];
    setCompanies(comps);
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
        (agData || [] as any[]).forEach((a: any) => {
          if (byAgent[String(a.id)]) a.linked_company_ids = byAgent[String(a.id)];
        });
      }
    } catch {
      // agent_companies table not created yet — single company_id fallback applies.
    }
    setAllAgents(agData || []);

    const nourish = comps.find((c) =>
      c.company_name?.toLowerCase().includes("nourish")
    );
    if (nourish) {
      setSelectedCompany(String(nourish.id));
    } else if (comps.length > 0) {
      setSelectedCompany(String(comps[0].id));
    }

    await loadVetVisits(uUpazilas, uDistricts);
    setLoading(false);
  }

  async function loadVaccinesList() {
    const { data } = await supabase
      .from("vaccines")
      .select("*")
      .order("age_days", { ascending: true });
    setVaccines(data || []);
  }

  // Insert with shed_id when the visit/record belongs to a real (non-legacy)
  // shed; if the column does not exist yet (pre-v2 DB), retry without it so
  // saving never breaks. Returns the Supabase result ({ error, ... }).
  async function insertWithOptionalShedId(
    table:
      | "vet_visits"
      | "before_sale_reports"
      | "after_sale_reports"
      | "layer_reports"
      | "farmer_feed_logs",
    payload: any,
    shedId: any
  ) {
    const numericShedId =
      shedId !== null &&
      shedId !== undefined &&
      String(shedId) !== "" &&
      !String(shedId).startsWith("legacy-")
        ? shedId
        : null;
    // Ownership (rule 2): stamp the creator on new rows. If the
    // created_by column does not exist yet, retry without it so the
    // save itself never fails for a missing column.
    const withCreator =
      currentUserId && (payload as any).created_by == null
        ? { ...payload, created_by: currentUserId }
        : payload;
    const doInsert = (p: any, withShed: boolean) => {
      if (withShed && numericShedId != null) {
        return supabase.from(table).insert([{ ...p, shed_id: numericShedId }]);
      }
      return supabase.from(table).insert([p]);
    };
    let res = await doInsert(withCreator, true);
    let strippedCreator = false;
    if (res.error && /created_by/i.test(String(res.error.message || ""))) {
      strippedCreator = true;
      res = await doInsert(payload, true);
    }
    if (res.error && /shed_id/i.test(String(res.error.message || ""))) {
      res = await doInsert(strippedCreator ? payload : withCreator, false);
    }
    return res;
  }

  async function updateWithOptionalShedId(
    table: "vet_visits",
    updates: Record<string, any>,
    shedId: any,
    eqColumn: string,
    eqValue: any
  ) {
    const numericShedId =
      shedId !== null &&
      shedId !== undefined &&
      String(shedId) !== "" &&
      !String(shedId).startsWith("legacy-")
        ? shedId
        : null;
    // Ownership (rule 2): stamp created_by on first edit of an old
    // row; never overwrite an existing creator. Best-effort: a
    // missing created_by column must not break the update.
    let withCreator: Record<string, any> = { ...updates };
    if (currentUserId) {
      try {
        const { data: existing } = await supabase
          .from(table)
          .select("created_by")
          .eq(eqColumn, eqValue)
          .maybeSingle();
        if (existing && !(existing as any).created_by) {
          withCreator = { ...updates, created_by: currentUserId };
        }
      } catch {
        // created_by column not added yet.
      }
    }
    const doUpdate = (u: Record<string, any>, withShed: boolean) => {
      if (withShed && numericShedId != null) {
        return supabase
          .from(table)
          .update({ ...u, shed_id: numericShedId })
          .eq(eqColumn, eqValue);
      }
      return supabase.from(table).update(u).eq(eqColumn, eqValue);
    };
    let res = await doUpdate(withCreator, true);
    let strippedCreator = false;
    if (res.error && /created_by/i.test(String(res.error.message || ""))) {
      strippedCreator = true;
      res = await doUpdate(updates, true);
    }
    if (res.error && /shed_id/i.test(String(res.error.message || ""))) {
      res = await doUpdate(strippedCreator ? updates : withCreator, false);
    }
    return res;
  }

  // Upsert a farmer's sheds into farmer_sheds (delete sheds removed from
  // the form). The farmer row itself is already saved by the caller, so a
  // missing farmer_sheds table only warns — it never loses the farmer.
  async function saveShedsForFarmer(
    farmerId: number,
    farmerName: string,
    sheds: any[],
    creatorId?: string | null
  ): Promise<boolean> {
    const list =
      sheds && sheds.length > 0
        ? sheds
        : [{ ...EMPTY_SHED_FORM, shed_name: "" }];
    try {
      const { data: existingSheds, error: existingErr } = await supabase
        .from("farmer_sheds")
        .select("id")
        .eq("farmer_id", farmerId);
      if (existingErr) throw existingErr;
      const keepIds = new Set(
        list
          .map((s: any) => s.id)
          .filter(
            (id: any) => id != null && !String(id).startsWith("legacy-")
          )
          .map((id: any) => Number(id))
      );
      const staleIds = (existingSheds || [])
        .map((s: any) => Number(s.id))
        .filter((id: number) => !keepIds.has(id));
      if (staleIds.length > 0) {
        await supabase.from("farmer_sheds").delete().in("id", staleIds);
      }
      for (let i = 0; i < list.length; i++) {
        const s = list[i];
        const db = shedFormToDb(s, farmerId, farmerName, i, list.length);
        const numericId =
          s.id != null && !String(s.id).startsWith("legacy-")
            ? Number(s.id)
            : null;
        if (numericId) {
          const { error } = await supabase
            .from("farmer_sheds")
            .update(db)
            .eq("id", numericId);
          if (error) throw error;
        } else {
          const { error } = await supabase
            .from("farmer_sheds")
            .insert([db]);
          if (error) throw error;
        }
      }
      // Ownership: stamp sheds that still have no creator with the
      // farmer's creator. Never overwrites a non-null created_by.
      try {
        const shedCreator = creatorId || currentUserId || null;
        if (shedCreator) {
          await supabase
            .from("farmer_sheds")
            .update({ created_by: shedCreator })
            .eq("farmer_id", farmerId)
            .is("created_by", null);
        }
      } catch {
        // created_by column not added yet on farmer_sheds.
      }
      return true;
    } catch (shedErr: any) {
      alert(
        "Farmer saved, but sheds could not be saved. Please run the farmer_sheds migration in Supabase first, then edit the farmer again to add sheds."
      );
      return false;
    }
  }

  // Apply a stock/category update to one shed's farmer_sheds row. Silently
  // skipped for legacy (synthesised) sheds and when the table/column is
  // missing — the caller still updates the farmer-row legacy totals.
  async function updateFarmerShedStock(
    shed: any,
    stockUpdates: Record<string, any>
  ): Promise<void> {
    try {
      if (!shed || shed.id == null || String(shed.id).startsWith("legacy-"))
        return;
      const { error } = await supabase
        .from("farmer_sheds")
        .update(stockUpdates)
        .eq("id", shed.id);
      if (error) throw error;
    } catch {
      // farmer_sheds table/column not created yet — legacy totals cover it.
    }
  }

  // The Add/Edit Farmer modal edits the FIRST shed through farmerForm; extra
  // sheds live in formSheds. Merge them into save-ready sheds: shed names
  // are normalised by getShedShortName() — a single shed stays blank (shown
  // as just the farmer name); with 2+ sheds a blank label becomes "Shed N"
  // (the full "Farmer Name - Shed N" is composed at display time by
  // getShedDisplayName).
  function syncFormShedsWithFarmerForm(): any[] {
    const base =
      formSheds && formSheds.length > 0
        ? formSheds
        : [{ ...EMPTY_SHED_FORM }];
    const farmerName = farmerForm.farmer_name.trim() || "Farmer";
    return base.map((s: any, i: number) => {
      const shed: any = { ...s };
      if (i === 0) {
        shed.shed_name = getShedShortName(
          { farmer_name: farmerName },
          s,
          0,
          base.length
        );
        shed.farm_category = farmerForm.farm_category || "";
        shed.latitude = farmerForm.latitude || "";
        shed.longitude = farmerForm.longitude || "";
        shed.feed_type = farmerForm.feed_type || "";
        shed.status = farmerForm.status || "Active";
        SHED_STOCK_FIELDS.forEach((field) => {
          shed[field] = Number((farmerForm as any)[field] || 0);
        });
      } else {
        shed.farm_category =
          shed.farm_category || farmerForm.farm_category || "";
        shed.shed_name = getShedShortName(
          { farmer_name: farmerName },
          s,
          i,
          base.length
        );
        if (shed.latitude === undefined) shed.latitude = "";
        if (shed.longitude === undefined) shed.longitude = "";
        SHED_STOCK_FIELDS.forEach((field) => {
          shed[field] = Number(shed[field] || 0);
        });
      }
      return shed;
    });
  }

  function updateFormShedField(index: number, field: string, value: any) {
    setFormSheds((prev: any[]) =>
      prev.map((s: any, i: number) => (i === index ? { ...s, [field]: value } : s))
    );
  }

  function addFormShed() {
    // Adding a 2nd shed: if the only existing shed has no name yet, ask for
    // its name first. Empty/cancel => it becomes "Shed 1" automatically.
    // The previous shed's name is saved to the DB right away (when it is an
    // existing shed row), not just in the form state.
    const prev = formSheds && formSheds.length > 0 ? formSheds[0] : null;
    const prevNeedsName =
      formSheds.length === 1 &&
      prev &&
      getShedShortName({ farmer_name: farmerForm.farmer_name }, prev, 0, 1) ===
        "";
    let prevName = "";
    if (prevNeedsName) {
      const answer = window.prompt(
        "Name the previous shed (leave empty for 'Shed 1'):",
        ""
      );
      prevName =
        answer && answer.trim()
          ? getShedShortName(
              { farmer_name: farmerForm.farmer_name },
              { shed_name: answer.trim() },
              0,
              2
            )
          : "Shed 1";
      if (prev && prev.id != null && !String(prev.id).startsWith("legacy-")) {
        const prevId = Number(prev.id);
        supabase
          .from("farmer_sheds")
          .update({ shed_name: prevName })
          .eq("id", prevId)
          .then(({ error }: any) => {
            if (error)
              console.error("Could not save previous shed name:", error);
          });
      }
    }
    // Prefill the SHORT label only ("Shed 2"); the full name
    // ("Farmer Name - Shed 2") is composed at display time.
    setFormSheds((prevSheds: any[]) => {
      const newCount = prevSheds.length + 1;
      const updated = prevNeedsName
        ? prevSheds.map((s: any, i: number) =>
            i === 0 ? { ...s, shed_name: prevName } : s
          )
        : prevSheds;
      return [
        ...updated,
        {
          ...EMPTY_SHED_FORM,
          shed_name: `Shed ${newCount}`,
          farm_category: farmerForm.farm_category || "Poultry",
        },
      ];
    });
  }

  function removeFormShed(index: number) {
    setFormSheds((prev: any[]) =>
      prev.length <= 1 ? prev : prev.filter((_: any, i: number) => i !== index)
    );
  }


  async function loadVetVisits(
    uUpa = assignedUpazilas,
    uDist = assignedDistricts
  ) {
    setLoading(true);
    try {
      const [
        { data: visits, error: visitErr },
        { data: agentsData },
        { data: farmersData },
        { data: bMeta },
        { data: aMeta },
        { data: lMeta },
        { data: subAgentsData },
      ] = await Promise.all([
        supabase
          .from("vet_visits")
          .select("*")
          .order("created_at", { ascending: false }),
        supabase.from("agents").select("*").order("agent_name"),
        supabase.from("farmers").select("*").order("farmer_name"),
        supabase
          .from("before_sale_reports")
          .select("id, vet_visit_id, farmer_id, visit_date"),
        supabase
          .from("after_sale_reports")
          .select("id, vet_visit_id, farmer_id, visit_date"),
        supabase
          .from("layer_reports")
          .select("id, vet_visit_id, farmer_id, visit_date"),
        // sub_agents table may not exist yet (migration not run) —
        // data is null then and the list stays empty.
        supabase
          .from("sub_agents")
          .select("id, sub_agent_name, agent_id, latitude, longitude"),
      ]);
      if (visitErr) throw visitErr;

      const vMetaMap: Record<string, { count: number; lastDate: string }> = {};
      const fMetaMap: Record<string, { count: number; lastDate: string }> = {};

      // Only reports attached to a LIVE vet_visits row count. A report
      // whose batch row was deleted (directly in Supabase) is an orphan:
      // the History modal can only show reports of live batches, so
      // counting orphans here enabled the History button + "Last:" date
      // for a modal that opens empty.
      const liveVisitIds = new Set<string>(
        (visits || []).map((v: any) => String(v.id))
      );

      const allReportRows = [
        ...(bMeta || []),
        ...(aMeta || []),
        ...(lMeta || []),
      ];
      allReportRows.forEach((r: any) => {
        const vKey = String(r.vet_visit_id || "");
        if (!vKey || !liveVisitIds.has(vKey)) return; // orphan — never shown
        const fKey = String(r.farmer_id || "");
        const dStr = String(r.visit_date || "").slice(0, 10);

        if (vKey) {
          if (!vMetaMap[vKey]) vMetaMap[vKey] = { count: 0, lastDate: "" };
          vMetaMap[vKey].count += 1;
          if (
            dStr &&
            (!vMetaMap[vKey].lastDate || dStr > vMetaMap[vKey].lastDate)
          ) {
            vMetaMap[vKey].lastDate = dStr;
          }
        }
        if (fKey) {
          if (!fMetaMap[fKey]) fMetaMap[fKey] = { count: 0, lastDate: "" };
          fMetaMap[fKey].count += 1;
          if (
            dStr &&
            (!fMetaMap[fKey].lastDate || dStr > fMetaMap[fKey].lastDate)
          ) {
            fMetaMap[fKey].lastDate = dStr;
          }
        }
      });

      setVisitReportMeta(vMetaMap);
      setFarmerReportMeta(fMetaMap);
      setSubAgents((subAgentsData as any[]) || []);

      const allVisits = visits || [];
      const agentsList = agentsData || [];
      let farmersList = farmersData || [];

      // Multi-shed: attach farmer_sheds rows when the table exists; the
      // legacy fallback in getShedsForFarmer() covers farmers without rows.
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
          farmersList.forEach((f: any) => {
            if (shedsByFarmer[String(f.id)])
              f.farmer_sheds = shedsByFarmer[String(f.id)];
          });
        }
      } catch {
        // farmer_sheds table not created yet — legacy fallback applies.
      }

      if (!isFullAccess) {
        farmersList = farmersList.filter((f: any) => {
          const matchUpa = uUpa.length === 0 || uUpa.includes(f.upazila);
          const matchDist = uDist.length === 0 || uDist.includes(f.district);
          return matchUpa && matchDist;
        });
      }

      const enrich = (v: any) => {
        const vFarmer = farmersList.find(
          (f: any) => String(f.id) === String(v.farmer_id)
        );
        const vSheds = vFarmer ? getShedsForFarmer(vFarmer) : [];
        const vShed = getShedForVisit(v, vFarmer);
        return {
          ...v,
          agents:
            agentsList.find((a: any) => String(a.id) === String(v.agent_id)) ||
            {
              agent_name: "Unknown Agent",
              upazila: "",
            },
          farmers: vFarmer || {
            farmer_name: "Unknown Farmer",
            mobile: "-",
            village: "-",
            upazila: "-",
          },
          shed: vShed,
          shedDisplayName: vShed
            ? getShedDisplayName(
                vFarmer,
                vShed,
                vSheds.indexOf(vShed),
                vSheds.length
              )
            : vFarmer?.farmer_name || "Unknown Farmer",
        };
      };

      const processed = farmersList.map((f: any) => {
        const fv = allVisits
          .filter((v: any) => String(v.farmer_id) === String(f.id))
          .sort((a: any, b: any) =>
            String(a.created_at || "").localeCompare(String(b.created_at || ""))
          );

        // Batch numbers count per SHED, not per farmer: each shed's
        // batches start again at 1 (Bridge par: Batch 1..N, Shed 2:
        // Batch 1..N). A visit without a shed_id on a multi-shed
        // farmer counts in its own group ("none").
        const perShedCount: Record<string, number> = {};
        const numberedVisits = fv.map((v: any) => {
          const enriched = enrich(v);
          const shedKey = String(enriched.shed?.id || "none");
          perShedCount[shedKey] = (perShedCount[shedKey] || 0) + 1;
          const seqNum = perShedCount[shedKey];
          const autoBatchTitle = `${f.farmer_name} Batch ${seqNum}`;
          return {
            ...enriched,
            seqNum,
            autoBatchTitle,
            displayBatchName: v.batch_label || autoBatchTitle,
          };
        });

        const runningVisits = numberedVisits.filter(
          (v: any) => v.status === "Activated" || v.status === "Running"
        );
        const closedVisits = [...numberedVisits]
          .filter((v: any) => v.status === "Sold" || v.status === "Culled")
          .reverse();
        const latestClosed = closedVisits[0] || null;

        // Sale date comes ONLY from a live closed visit row (Sold/Culled).
        // farmers.last_sale_date is a stale copy: it survives when the
        // sale record is deleted (in the app or directly in Supabase),
        // so it must never drive this display.
        const lastSaleDate =
          latestClosed?.closed_date ||
          (latestClosed?.created_at
            ? String(latestClosed.created_at).slice(0, 10)
            : null);
        const daysSinceSale = getDaysSince(lastSaleDate);

        let computedStatus: "Running" | "Active" | "Inactive" = "Inactive";
        let statusDate = "";
        if (runningVisits.length > 0) {
          computedStatus = "Running";
          statusDate = runningVisits[0]?.hatch_date || "";
        } else if (daysSinceSale !== null && daysSinceSale <= 60) {
          computedStatus = "Active";
          statusDate = lastSaleDate || "";
        } else {
          computedStatus = "Inactive";
          statusDate = lastSaleDate || "";
        }

        const agentObj = agentsList.find(
          (a: any) => String(a.id) === String(f.agent_id)
        );
        // Per-shed status (same rule as computedStatus, keyed by shed)
        // so the table can show Active/Inactive sheds too, not only
        // the sheds that currently have a Running batch.
        const shedStatuses = getShedsForFarmer(f).map((shed: any) => {
          const st = getShedVisitStatus(f, shed, fv);
          return { shed, ...st };
        });
        return {
          ...f,
          computedStatus,
          statusDate,
          daysSinceSale,
          shedStatuses,
          closedType: latestClosed?.status || "",
          closedCategory: latestClosed?.category || "",
          category: runningVisits[0]?.category || normalizeCategory(f),
          allBatches: [...numberedVisits].reverse(),
          runningVisits,
          runningCount: runningVisits.length,
          latestClosed,
          agentName: agentObj?.agent_name || "Unassigned Agent",
          agentUpazila: agentObj?.upazila || f.upazila || "",
          nourishTag: f.nourish_status || "",
          nourishTagDate: f.nourish_status_date || "",
        };
      });

      setStatusFarmers(processed);
      setVisitsData(
        allVisits
          .filter((v: any) => v.status === "Activated" || v.status === "Running")
          .map(enrich)
      );
    } catch (err) {
      console.error("Error loading vet visits:", err);
    }
    setLoading(false);
  }

  function getAgeDays(hatchDateStr?: string, refDateStr?: string): number {
    if (!hatchDateStr) return 0;
    const hatchDate = new Date(hatchDateStr);
    const refDate = refDateStr ? new Date(refDateStr) : new Date();
    const diffTime = refDate.getTime() - hatchDate.getTime();
    return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
  }

  function getAgeWeeks(hatchDateStr?: string, refDateStr?: string): number {
    const days = getAgeDays(hatchDateStr, refDateStr);
    return Math.max(1, Math.floor(days / 7));
  }

  function calculateAge(
    hatchDateStr: string,
    category: string,
    refDateStr?: string
  ) {
    if (!hatchDateStr) return "N/A";
    const diffDays = getAgeDays(hatchDateStr, refDateStr);
    const catLower = (category || "").toLowerCase();
    if (catLower.includes("layer")) {
      const totalWeeks = (diffDays / 7).toFixed(1);
      return `${totalWeeks}W (${diffDays}d)`;
    }
    return `${diffDays}d`;
  }

  function getMeatStdForVisit(categoryName?: string, ageDays?: number) {
    const d = Number(ageDays || 0);
    const cat = String(categoryName || "Broiler")
      .trim()
      .toLowerCase();
    if (cat === "sonali") {
      return sonaliStds[d] || null;
    }
    if (cat === "broiler") {
      return broilerStds[d] || null;
    }
    return null;
  }

  // Accurate Layer Standard Lookup for White vs Brown
  function getLayerStdForVisit(breedName?: string, ageWks?: number) {
    const w = Math.min(100, Math.max(1, ageWks || 1));
    const bLower = String(breedName || "")
      .trim()
      .toLowerCase();
    const isWhite = isWhiteLayerBreedName(bLower);

    if (bLower && lMapHasExact(layerStds, `${bLower}_${w}`)) {
      return layerStds[`${bLower}_${w}`];
    }
    if (isWhite) {
      return (
        layerStds[`white_${w}`] ||
        layerStds[`bovans white_${w}`] ||
        layerStds[`any_${w}`] ||
        null
      );
    }
    return (
      layerStds[`brown_${w}`] ||
      layerStds[`shaver star cross 579 brown_${w}`] ||
      layerStds[`shaver star cross 579_${w}`] ||
      layerStds[`any_${w}`] ||
      null
    );
  }

  function lMapHasExact(map: Record<string, any>, key: string): boolean {
    return Boolean(map && map[key]);
  }

  function getVaccineInfoForVisit(visit: any) {
    const cat = (visit.category || "Broiler").trim();
    const ageDays = getAgeDays(visit.hatch_date);
    const catVaccines = vaccines
      .filter(
        (v) =>
          v.category?.trim().toLowerCase() === cat.toLowerCase() &&
          Number(v.age_days) > 0
      )
      .sort((a, b) => Number(a.age_days) - Number(b.age_days));

    if (catVaccines.length === 0)
      return { upcoming: null, upcomingDateStr: "", catVaccines: [] };

    const lastGivenName = String(visit.last_vaccine_name || "")
      .trim()
      .toLowerCase();

    const lastGivenObj = catVaccines.find((v) =>
      lastGivenName.includes(v.vaccine_name?.trim().toLowerCase())
    );
    const minNextAge = lastGivenObj
      ? Math.max(ageDays, Number(lastGivenObj.age_days))
      : ageDays;

    const upcoming = catVaccines.find(
      (v) =>
        Number(v.age_days) > minNextAge &&
        !lastGivenName.includes(v.vaccine_name?.trim().toLowerCase())
    );

    let upcomingDateStr = "";
    if (upcoming && visit.hatch_date) {
      const d = new Date(visit.hatch_date);
      d.setDate(d.getDate() + Number(upcoming.age_days));
      upcomingDateStr = d.toISOString().slice(0, 10);
    }

    return { upcoming, upcomingDateStr, catVaccines };
  }

  function handleOpenLogVaccineModal(visit: any) {
    const info = getVaccineInfoForVisit(visit);
    const defaultVacName =
      info.upcoming?.vaccine_name || visit.last_vaccine_name || "";

    setActiveVaccineVisit(visit);
    setLogVaccineForm({
      vaccine_name: defaultVacName,
      vaccine_date: new Date().toISOString().split("T")[0],
    });
    setShowLogVaccineModal(true);
  }

  async function handleSaveLoggedVaccine(e: React.FormEvent) {
    e.preventDefault();
    if (savingVaccine) return;
    setSavingVaccine(true);
    try {
    if (!activeVaccineVisit || !logVaccineForm.vaccine_name.trim()) return;
    // Logging a vaccine edits the visit row: Poultry category +
    // ownership.
    if (!canEditPoultryReport(activeVaccineVisit)) {
      alert("You can only log vaccines on visits you added.");
      return;
    }
    // Vaccine log date: never after today, never before hatch date.
    const vacDateErr = validateReportDate(
      logVaccineForm.vaccine_date,
      activeVaccineVisit.hatch_date
    );
    if (vacDateErr) {
      alert(vacDateErr);
      return;
    }

    try {
      const { error } = await supabase
        .from("vet_visits")
        .update({
          last_vaccine_name: logVaccineForm.vaccine_name.trim(),
          last_vaccine_date: logVaccineForm.vaccine_date,
        })
        .eq("id", activeVaccineVisit.id);

      if (error) throw error;

      // Ownership (rule 2): stamp created_by on first edit of an old row.
      if (!activeVaccineVisit.created_by && currentUserId) {
        try {
          await supabase
            .from("vet_visits")
            .update({ created_by: currentUserId })
            .eq("id", activeVaccineVisit.id);
        } catch {
          // created_by column not added yet.
        }
      }

      setShowLogVaccineModal(false);
      setActiveVaccineVisit(null);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save vaccine log.");
    }
    } finally {
      setSavingVaccine(false);
    }
  }

  function updateBeforeCalculations(updatedFields: any) {
    const tBirds = Number(activeReportVisit?.t_birds) || 0;
    const feedTKg =
      Number(
        updatedFields.feed_t_kg !== undefined
          ? updatedFields.feed_t_kg
          : beforeForm.feed_t_kg
      ) || 0;
    const mortality =
      Number(
        updatedFields.mortality_pcs !== undefined
          ? updatedFields.mortality_pcs
          : beforeForm.mortality_pcs
      ) || 0;
    const weightGmB =
      Number(
        updatedFields.weight_gm_b !== undefined
          ? updatedFields.weight_gm_b
          : beforeForm.weight_gm_b
      ) || 0;

    let feedGmB = beforeForm.feed_gm_b;
    let fcrWOutM = "";
    let fcrWithM = "";

    const liveBirds = Math.max(0, tBirds - mortality);

    if (tBirds > 0 && feedTKg > 0) {
      feedGmB = ((feedTKg * 1000) / tBirds).toFixed(1);
    } else if (updatedFields.feed_t_kg === "") {
      feedGmB = "";
    }

    if (weightGmB > 0 && feedTKg > 0) {
      const totalWeightWithoutMortalityKg = (tBirds * weightGmB) / 1000;
      const totalWeightWithMortalityKg = (liveBirds * weightGmB) / 1000;

      if (totalWeightWithoutMortalityKg > 0) {
        fcrWOutM = (feedTKg / totalWeightWithoutMortalityKg).toFixed(3);
      }
      if (totalWeightWithMortalityKg > 0) {
        fcrWithM = (feedTKg / totalWeightWithMortalityKg).toFixed(3);
      }
    }

    setBeforeForm((prev) => ({
      ...prev,
      ...updatedFields,
      feed_gm_b: feedGmB,
      fcr_w_out_m: fcrWOutM,
      fcr_with_m: fcrWithM,
    }));
  }

  function updateAfterCalculations(updatedFields: any) {
    const tBirds = Number(activeReportVisit?.t_birds) || 0;
    const feedKg =
      Number(
        updatedFields.feed_consumption_kg !== undefined
          ? updatedFields.feed_consumption_kg
          : afterForm.feed_consumption_kg
      ) || 0;
    const mortality =
      Number(
        updatedFields.mortality !== undefined
          ? updatedFields.mortality
          : afterForm.mortality
      ) || 0;
    const weightGm =
      Number(
        updatedFields.weight !== undefined
          ? updatedFields.weight
          : afterForm.weight
      ) || 0;

    let feedGmB = afterForm.feed_gm_b;
    let fcrWOutM = "";
    let fcrWithM = "";

    const liveBirds = Math.max(0, tBirds - mortality);

    if (tBirds > 0 && feedKg > 0) {
      feedGmB = ((feedKg * 1000) / tBirds).toFixed(1);
    } else if (updatedFields.feed_consumption_kg === "") {
      feedGmB = "";
    }

    if (weightGm > 0 && feedKg > 0) {
      const totalWeightWithoutMortalityKg = (tBirds * weightGm) / 1000;
      const totalWeightWithMortalityKg = (liveBirds * weightGm) / 1000;

      if (totalWeightWithoutMortalityKg > 0) {
        fcrWOutM = (feedKg / totalWeightWithoutMortalityKg).toFixed(3);
      }
      if (totalWeightWithMortalityKg > 0) {
        fcrWithM = (feedKg / totalWeightWithMortalityKg).toFixed(3);
      }
    }

    setAfterForm((prev) => ({
      ...prev,
      ...updatedFields,
      feed_gm_b: feedGmB,
      fcr_w_out_m: fcrWOutM,
      fcr_with_m: fcrWithM,
    }));
  }

  function handleLayerFeedChange(
    mode: "total_kg" | "gm_bird",
    rawValue: string,
    overrideMortality?: string
  ) {
    const baseBirds = Number(activeReportVisit?.t_birds) || 0;
    const mort = Number(
      overrideMortality !== undefined
        ? overrideMortality
        : layerForm.mortality_pcs || 0
    );
    // While editing, t_birds already had the old mortality subtracted, so
    // add it back before subtracting the new mortality (same recompute as
    // the save path); on add the old mortality is 0.
    const oldMortForFeed =
      editingReportId !== null && editingReportId !== undefined
        ? Math.max(
            0,
            Number(
              allReportsForFarmer.layer.find(
                (r: any) => String(r.id) === String(editingReportId)
              )?.mortality ??
                allReportsForFarmer.layer.find(
                  (r: any) => String(r.id) === String(editingReportId)
                )?.mortality_pcs ??
                0
            )
          )
        : 0;
    const liveBirds = Math.max(0, baseBirds + oldMortForFeed - mort);
    const num = Number(rawValue);

    if (mode === "total_kg") {
      const gmPerBird =
        liveBirds > 0 && num > 0 ? ((num * 1000) / liveBirds).toFixed(1) : "";
      setLayerForm((prev) => ({
        ...prev,
        feed_total_kg: rawValue,
        feed_intake_g: gmPerBird,
      }));
    } else {
      const totalKg =
        liveBirds > 0 && num > 0 ? ((num * liveBirds) / 1000).toFixed(2) : "";
      setLayerForm((prev) => ({
        ...prev,
        feed_intake_g: rawValue,
        feed_total_kg: totalKg,
      }));
    }
  }

  // Report validation (same rules on every report submit path):
  // a report date can never be after today or before the batch's hatch
  // date, and mortality can never exceed total birds. Returns the exact
  // Banglish alert message, or null when the value is allowed.
  function validateReportDate(
    reportDate: string,
    hatchDate: string | null | undefined
  ): string | null {
    const today = new Date().toISOString().split("T")[0];
    if (reportDate && reportDate > today) {
      return "You cannot add a report for a future date.";
    }
    if (reportDate && hatchDate && reportDate < hatchDate) {
      return "Report date cannot be before the hatch date.";
    }
    return null;
  }

  function validateMortality(mortality: number, cap: number): string | null {
    if (Number(mortality) > Number(cap)) {
      return "Mortality cannot be more than total birds.";
    }
    return null;
  }

  function validateBeforeSaleInputs(formObj: any): boolean {
    if (
      String(formObj.mortality_pcs ?? "").trim() === "" ||
      !Number(formObj.weight_gm_b) ||
      !Number(formObj.feed_t_kg) ||
      !String(formObj.feed_company || "").trim() ||
      !String(formObj.feed_type || "").trim()
    ) {
      alert(
        "Mortality, Weight, Total Feed, Feed Company, and Feed Type are mandatory!"
      );
      return false;
    }
    return true;
  }

  function validateAfterSaleInputs(formObj: any): boolean {
    if (
      String(formObj.mortality ?? "").trim() === "" ||
      !Number(formObj.weight) ||
      !Number(formObj.feed_consumption_kg) ||
      !String(formObj.feed_mill || "").trim() ||
      !String(formObj.feed || "").trim()
    ) {
      alert(
        "Total Mortality, Avg Weight, Total Feed Consumption, Feed Company, and Last Feed are mandatory!"
      );
      return false;
    }
    return true;
  }

  function validateLayerInputs(formObj: any): boolean {
    if (
      String(formObj.mortality_pcs ?? "").trim() === "" ||
      !Number(formObj.b_weight_g) ||
      (!Number(formObj.feed_intake_g) && !Number(formObj.feed_total_kg)) ||
      String(formObj.egg_production ?? "").trim() === "" ||
      String(formObj.egg_weight ?? "").trim() === "" ||
      !String(formObj.feed_company || "").trim() ||
      !String(formObj.feed_type || "").trim()
    ) {
      alert(
        "Mortality, Body Weight, Feed Intake, Egg Production, Egg Weight, Feed Company, and Feed Type are mandatory!"
      );
      return false;
    }
    return true;
  }

  // Header Builder: Handles both Nourish vs Non-Nourish for Broiler, Sonali & Layer
  function buildWhatsAppHeaderLine(
    hatcheryName: string,
    feedCompName: string,
    categoryName: string,
    defaultTitle: string
  ): string {
    const hClean = String(hatcheryName || "Nourish").trim();
    const fClean = String(feedCompName || "Nourish").trim();
    const bothNourish =
      hClean.toLowerCase().includes("nourish") &&
      fClean.toLowerCase().includes("nourish");

    if (bothNourish) {
      return `*${defaultTitle}*`;
    }
    const catLabel = categoryName || "Broiler";
    if (catLabel.toLowerCase() === "layer") {
      return `*${hClean} Layer chick and ${fClean} layer feed performance report:*`;
    }
    return `*${hClean} ${catLabel} chick and ${fClean} ${catLabel} feed Performance Report:*`;
  }

  // WhatsApp Agent line: a farmer under a sub-agent shows the sub-agent
  // ("sub-agent : Name") instead of the mother agent; direct-agent
  // farmers show the agent name as before. visit.farmers is the full
  // farmer row (farmers select "*"), so sub_agent_id is available.
  function getWhatsAppAgentLine(farmer: any, agent: any): string {
    const subId = farmer?.sub_agent_id ? String(farmer.sub_agent_id) : "";
    const subName = subId ? subAgentById[subId]?.sub_agent_name : "";
    return subName ? `sub-agent : ${subName}` : agent?.agent_name || "-";
  }

  function generateBroilerWhatsAppText(
    visit: any,
    rep: any,
    vDate: string,
    isFromForm = false
  ) {
    if (isFromForm && !validateBeforeSaleInputs(rep)) return;

    const farmer = visit.farmers || {};
    const agent = visit.agents || {};
    const catName = visit.category || "Broiler";
    const ageDays = getAgeDays(visit.hatch_date, vDate);
    const std = getMeatStdForVisit(catName, ageDays);
    const hasStd = Boolean(std && Number(std.weight) > 0);

    const actualWt = Number(rep.weight_gm_b || 0);
    const stdWt = Number(std?.weight || 0);
    const wtDiff = actualWt - stdWt;
    const wtDiffText =
      hasStd && stdWt > 0
        ? `${Math.abs(wtDiff).toFixed(0)}gm ${wtDiff >= 0 ? "more" : "less"}`
        : "";

    const actualFeed = Number(rep.feed_gm_b || 0);
    const stdFeed = Number(std?.feed_intake || 0);
    const feedDiff = actualFeed - stdFeed;
    const feedDiffText =
      hasStd && stdFeed > 0
        ? `${Math.abs(feedDiff).toFixed(0)}gm ${
            feedDiff >= 0 ? "more" : "less"
          }`
        : "";

    const findingsList: string[] = [];
    if (rep.indigestion && String(rep.indigestion).trim() !== "") {
      findingsList.push(
        `${findingsList.length + 1}. Indigestion (%) - ${rep.indigestion}%`
      );
    }
    if (rep.orange_feces && String(rep.orange_feces).trim() !== "") {
      findingsList.push(
        `${findingsList.length + 1}. Orange Faeces (%) - ${rep.orange_feces}%`
      );
    }
    if (rep.watery_feces && String(rep.watery_feces).trim() !== "") {
      findingsList.push(
        `${findingsList.length + 1}. Watery Faeces (%) - ${rep.watery_feces}%`
      );
    }
    if (rep.uneven && String(rep.uneven).trim() !== "") {
      findingsList.push(
        `${findingsList.length + 1}. Uneven (%) - ${rep.uneven}%`
      );
    }

    const findingsBlock =
      findingsList.length > 0
        ? `--------------------------------------------\nFindings :\n${findingsList.join(
            "\n"
          )}\n`
        : "";

    const rmkText = getRemarkText(rep);
    const remarksBlock = rmkText
      ? `--------------------------------------------\nRemarks: ${rmkText}\n`
      : "";

    const headerTitle = buildWhatsAppHeaderLine(
      visit.hatchery || "Nourish",
      rep.feed_company || "Nourish",
      catName,
      `${catName} Farm visit Observation Report`
    );

    const actualFarmerName =
      getVisitShedDisplayName(visit, farmer) || visit.batch_label || "-";

    const weightSection = hasStd
      ? `Body Weight - ${actualWt} gm/bird.\nStd - ${stdWt} gm/bird.`
      : `Body Weight - ${actualWt} gm/bird.`;

    const feedSection = hasStd
      ? `Feed Intake - ${actualFeed} gm/bird.\nStd - ${stdFeed} gm/bird.`
      : `Feed Intake - ${actualFeed} gm/bird.`;

    const fcrSection = hasStd
      ? `FCR (without mortality): ${rep.fcr_w_out_m || "-"}\nStd - ${
          std?.fcr || "-"
        }\nFCR (with mortality): ${rep.fcr_with_m || "-"}`
      : `FCR (without mortality): ${
          rep.fcr_w_out_m || "-"
        }\nFCR (with mortality): ${rep.fcr_with_m || "-"}`;

    const summaryLine =
      hasStd && wtDiffText && feedDiffText
        ? `\n*Body weight (${wtDiffText}) than standard and Feed intake (${feedDiffText}) than standard*\n`
        : "";

    const msg = `Assalamu Alaikum sir.
${headerTitle}
Visit Date: ${formatDmy(vDate)}

Agent: ${getWhatsAppAgentLine(farmer, agent)}
Farmer Name - ${actualFarmerName}
Farm Address - ${[farmer.village, farmer.upazila].filter(Boolean).join(", ")}
--------------------------------------------
Hatchery - ${visit.hatchery || "-"}
Breed - ${visit.breed || "-"}
H.Date: ${formatDmy(visit.hatch_date)}
--------------------------------------------
Total Bird (Pcs) - ${visit.t_birds || 0}
Age (Day) - ${ageDays} days
Mortality (Pcs) - ${rep.mortality_pcs ?? 0} pcs

${weightSection}

${feedSection}

${fcrSection}
--------------------------------------------
Feed - ${rep.feed_company || "-"}
Feed Type - ${rep.feed_type || "-"}
MFG: ${rep.pro_date ? formatDmy(rep.pro_date) : "-"}
Batch: ${rep.batch_no || "-"}
${findingsBlock}${remarksBlock}${summaryLine}
Thanks.
${currentUserProfile.fullName}
${currentUserProfile.designation}
${currentUserProfile.region}`;

    setCopiedWa(false);
    setWhatsappPopupText(msg);
  }

  // Layer WhatsApp Message (Resolves Breed accurately from Batch & Report + White/Brown Standard + Non-Nourish Header)
  function generateLayerWhatsAppText(
    visit: any,
    rep: any,
    isFromForm = false
  ) {
    if (isFromForm && !validateLayerInputs(rep)) return;

    const farmer = visit.farmers || {};
    const agent = visit.agents || {};
    const vDate = rep.visit_date || reportVisitDate;
    const ageWks = getAgeWeeks(visit.hatch_date, vDate);

    // Priority: visit.breed (batch breed) -> rep.breed -> fallback
    const breedName =
      (isFromForm
        ? rep.breed || visit.breed
        : visit.breed || rep.breed) || "Shaver Star Cross 579 Brown";

    const std = getLayerStdForVisit(breedName, ageWks);

    const hatcheryName =
      (isFromForm
        ? rep.hatchery || visit.hatchery
        : visit.hatchery || rep.hatchery) || "Nourish";
    const feedCompName = rep.feed_company || "Nourish";

    const headerTitle = buildWhatsAppHeaderLine(
      hatcheryName,
      feedCompName,
      "Layer",
      "Layer farm visit report"
    );

    const actualFarmerName =
      getVisitShedDisplayName(visit, farmer) || visit.batch_label || "-";

    const mortVal = Number(rep.mortality ?? rep.mortality_pcs ?? 0);
    const baseBirds = Number(visit.t_birds || 0);
    // From the form, the live total is recomputed from scratch: on edit the
    // visit's t_birds already had the old mortality subtracted, so add it
    // back before subtracting the new mortality.
    const editingOldMortForWa =
      isFromForm && editingReportId !== null && editingReportId !== undefined
        ? Math.max(
            0,
            Number(
              allReportsForFarmer.layer.find(
                (r: any) => String(r.id) === String(editingReportId)
              )?.mortality ??
                allReportsForFarmer.layer.find(
                  (r: any) => String(r.id) === String(editingReportId)
                )?.mortality_pcs ??
                0
            )
          )
        : 0;
    const displayBirds = isFromForm
      ? Math.max(0, baseBirds + editingOldMortForWa - mortVal)
      : baseBirds;

    const rmkText = getRemarkText(rep);
    const remarksLine = rmkText ? `\nRemarks: ${rmkText}\n` : "";

    const msg = `Assalamu Alaikum sir.
${headerTitle}
Visit Date: ${formatDmy(vDate)}

Agent Name: ${getWhatsAppAgentLine(farmer, agent)}
Farmer: ${actualFarmerName}
Address: ${[farmer.village, farmer.upazila].filter(Boolean).join(", ")}

Hatchery: ${hatcheryName}
Breed: ${breedName}
H.Date: ${formatDmy(visit.hatch_date)}

T.B.: ${displayBirds}
Age: ${ageWks} weeks

Weight: ${rep.b_weight_g || 0} gm
Std Weight: ${std?.weight || "-"} gm

Feed Intake(gm/birds): ${rep.feed_intake_g || 0} gm
Std Feed: ${std?.feed_bird_day || "-"} gm

Eggs Production: ${rep.egg_production || 0}%
Std Egg Production: ${std?.egg_production || "-"}%

Egg weight: ${rep.egg_weight || 0} gm
Std Egg weight: ${std?.egg_weight || "-"} gm

Feed: ${rep.feed_company || "-"}
Type: ${rep.feed_type || "-"}
MFG: ${rep.pro_date ? formatDmy(rep.pro_date) : "-"}
Batch No.: ${rep.batch_no || "-"}
Disease: ${rep.disease || "-"}
${remarksLine}
Thanks.
${currentUserProfile.fullName}
${currentUserProfile.designation}
${currentUserProfile.region}`;

    setCopiedWa(false);
    setWhatsappPopupText(msg);
  }

  function getFeedChangeCompanyObj(feedCompanyName: string) {
    if (!feedCompanyName || !activeReportVisit) return null;
    const chosenComp = companies.find(
      (c) =>
        c.company_name?.trim().toLowerCase() ===
        feedCompanyName.trim().toLowerCase()
    );
    if (!chosenComp) return null;
    if (String(chosenComp.id) === String(activeReportVisit.company_id)) {
      return null;
    }
    return chosenComp;
  }

  async function processFeedCompanyChangeIfNeeded(
    visit: any,
    chosenFeedCompanyName: string,
    chosenNewAgentId: string,
    visitDateStr: string
  ) {
    if (!chosenFeedCompanyName) return;

    const currentCompObj = companies.find(
      (c) => String(c.id) === String(visit.company_id)
    );
    const newCompObj = companies.find(
      (c) =>
        c.company_name?.trim().toLowerCase() ===
        chosenFeedCompanyName.trim().toLowerCase()
    );

    const prevFeedName =
      visit.farmers?.last_feed_company || currentCompObj?.company_name || "";
    const wasNourish = prevFeedName.toLowerCase().includes("nourish");
    const isNowNourish = chosenFeedCompanyName
      .toLowerCase()
      .includes("nourish");

    const farmerUpdates: Record<string, any> = {
      last_feed_company: chosenFeedCompanyName,
    };

    let convType = "";
    if (wasNourish && !isNowNourish) {
      convType = "Leave from Nourish";
      farmerUpdates.nourish_status = convType;
      farmerUpdates.nourish_status_date = visitDateStr;
    } else if (!wasNourish && isNowNourish) {
      convType = "Convert to Nourish";
      farmerUpdates.nourish_status = convType;
      farmerUpdates.nourish_status_date = visitDateStr;
    }

    if (
      prevFeedName &&
      prevFeedName.trim().toLowerCase() !==
        chosenFeedCompanyName.trim().toLowerCase()
    ) {
      await insertWithOptionalShedId(
        "farmer_feed_logs",
        {
          farmer_id: String(visit.farmer_id),
          from_company: prevFeedName,
          to_company: chosenFeedCompanyName,
          conversion_type: convType || "Feed Switch",
          change_date: visitDateStr,
        },
        visit.shed_id
      );
    }

    if (newCompObj && String(newCompObj.id) !== String(visit.company_id)) {
      if (!chosenNewAgentId) {
        throw new Error(
          `Feed changed to ${newCompObj.company_name}. Please select an Agent under ${newCompObj.company_name}.`
        );
      }
      farmerUpdates.company_id = Number(newCompObj.id);
      farmerUpdates.agent_id = Number(chosenNewAgentId);

      await supabase
        .from("vet_visits")
        .update({
          company_id: String(newCompObj.id),
          agent_id: String(chosenNewAgentId),
        })
        .eq("id", visit.id);
    }

    await supabase
      .from("farmers")
      .update(farmerUpdates)
      .eq("id", visit.farmer_id);
  }

  function openReportModal(
    type: "before_sale" | "after_sale" | "layer_report",
    visit: any
  ) {
    if (!canAddPoultryReport) {
      alert("Your assigned categories do not include Poultry, so you cannot add reports here.");
      return;
    }
    setEditingReportId(null);
    setEditingReportCreatedBy(null);
    setActiveReportVisit(visit);
    setReportVisitDate(new Date().toISOString().split("T")[0]);
    setTransferAgentId("");
    setShowCullReasonInput(false);
    setCullReasonText("");

    const currentCompName =
      companies.find((c) => String(c.id) === String(visit.company_id))
        ?.company_name || "";
    const cat = visit.category || "Broiler";
    const defaultFeedTypes =
      UPDATED_FEED_TYPES_BY_CATEGORY[cat] ||
      UPDATED_FEED_TYPES_BY_CATEGORY["Broiler"];

    if (type === "before_sale") {
      setBeforeForm({
        mortality_pcs: "",
        weight_gm_b: "",
        feed_t_kg: "",
        feed_gm_b: "",
        fcr_w_out_m: "",
        fcr_with_m: "",
        uneven: "",
        indigestion: "",
        orange_feces: "",
        watery_feces: "",
        feeder: "",
        drinker: "",
        floor_space: "",
        litter_condition: "",
        ventilation: "",
        feed_scattering: "",
        feed_hardness: "",
        feed_company: currentCompName,
        feed_type: defaultFeedTypes[0] || "Broiler Pre Starter",
        pro_date: "",
        batch_no: "",
        remarks: "",
      });
      setShowBeforeSaleModal(true);
    } else if (type === "after_sale") {
      setAfterForm({
        mortality: "",
        weight: "",
        feed_consumption_kg: "",
        feed_gm_b: "",
        feed:
          defaultFeedTypes[defaultFeedTypes.length - 1] || "Broiler Finisher",
        feed_mill: currentCompName,
        fcr_w_out_m: "",
        fcr_with_m: "",
        remarks: "",
      });
      setShowAfterSaleModal(true);
    } else if (type === "layer_report") {
      setLayerForm({
        mortality_pcs: "0",
        b_weight_g: "",
        feed_total_kg: "",
        feed_intake_g: "",
        egg_production: "",
        egg_weight: "",
        feed_company: currentCompName,
        feed_type: "Layer Starter",
        pro_date: "",
        batch_no: "",
        hatchery: visit.hatchery || "",
        breed: visit.breed || "Shaver Star Cross 579 Brown",
        disease: "",
        remarks: "",
      });
      setShowLayerModal(true);
    }
  }

  // Open existing report in the exact same Report Modal (with Std values & Delete icon)
  function openEditReport(
    table: "before_sale_reports" | "after_sale_reports" | "layer_reports",
    rep: any
  ) {
    if (!canEditPoultryReport(rep)) {
      alert("You can only edit reports you added.");
      return;
    }
    const visit = selectedHistoryVisit;
    if (!visit) return;

    setEditingReportId(rep.id);
    setEditingReportCreatedBy(rep?.created_by ?? null);
    setActiveReportVisit(visit);
    setReportVisitDate(
      rep.visit_date || new Date().toISOString().split("T")[0]
    );
    setTransferAgentId("");
    setShowCullReasonInput(false);
    setCullReasonText("");

    if (table === "before_sale_reports") {
      setBeforeForm({
        mortality_pcs: String(rep.mortality_pcs ?? ""),
        weight_gm_b: String(rep.weight_gm_b ?? ""),
        feed_t_kg: String(rep.feed_t_kg ?? ""),
        feed_gm_b: String(rep.feed_gm_b ?? ""),
        fcr_w_out_m: String(rep.fcr_w_out_m ?? ""),
        fcr_with_m: String(rep.fcr_with_m ?? ""),
        uneven: rep.uneven || "",
        indigestion: rep.indigestion || "",
        orange_feces: rep.orange_feces || "",
        watery_feces: rep.watery_feces || "",
        feeder: rep.feeder || "",
        drinker: rep.drinker || "",
        floor_space: rep.floor_space || "",
        litter_condition: rep.litter_condition || "",
        ventilation: rep.ventilation || "",
        feed_scattering: rep.feed_scattering || "",
        feed_hardness: rep.feed_hardness || "",
        feed_company: rep.feed_company || "",
        feed_type: rep.feed_type || "Broiler Pre Starter",
        pro_date: rep.pro_date || "",
        batch_no: rep.batch_no || "",
        remarks: getRemarkText(rep),
      });
      setShowBeforeSaleModal(true);
    } else if (table === "after_sale_reports") {
      const tB = Number(visit.t_birds || 0);
      const calcGm =
        rep.feed_gm_b ||
        (tB > 0 && Number(rep.feed_consumption_kg) > 0
          ? ((Number(rep.feed_consumption_kg) * 1000) / tB).toFixed(1)
          : "");
      setAfterForm({
        mortality: String(rep.mortality ?? ""),
        weight: String(rep.weight ?? ""),
        feed_consumption_kg: String(rep.feed_consumption_kg ?? ""),
        feed_gm_b: String(calcGm ?? ""),
        feed: rep.feed || "Broiler Finisher",
        feed_mill: rep.feed_mill || "",
        fcr_w_out_m: String(rep.fcr_w_out_m ?? ""),
        fcr_with_m: String(rep.fcr_with_m ?? ""),
        remarks: getRemarkText(rep),
      });
      setShowAfterSaleModal(true);
    } else if (table === "layer_reports") {
      const baseBirds = Number(visit.t_birds || 0);
      const intakeG = Number(rep.feed_intake_g || 0);
      const calcTotalKg =
        rep.feed_total_kg ||
        (baseBirds > 0 && intakeG > 0
          ? ((intakeG * baseBirds) / 1000).toFixed(2)
          : "");
      setLayerForm({
        mortality_pcs: String(rep.mortality ?? rep.mortality_pcs ?? "0"),
        b_weight_g: String(rep.b_weight_g ?? ""),
        feed_total_kg: String(calcTotalKg ?? ""),
        feed_intake_g: String(rep.feed_intake_g ?? ""),
        egg_production: String(rep.egg_production ?? ""),
        egg_weight: String(rep.egg_weight ?? ""),
        feed_company: rep.feed_company || "",
        feed_type: rep.feed_type || "Layer Starter",
        pro_date: rep.pro_date || "",
        batch_no: rep.batch_no || "",
        hatchery: rep.hatchery || visit.hatchery || "",
        breed: rep.breed || visit.breed || "Shaver Star Cross 579 Brown",
        disease: rep.disease || "",
        remarks: getRemarkText(rep),
      });
      setShowLayerModal(true);
    }
  }

  async function handleDeleteCurrentReport(
    table: "before_sale_reports" | "after_sale_reports" | "layer_reports"
  ) {
    if (!editingReportId) return;
    // Permission guard (rule 5): Poultry report delete needs the
    // category assigned AND ownership.
    if (!canEditPoultryReport({ created_by: editingReportCreatedBy })) {
      alert("You can only delete reports you added.");
      return;
    }
    const ok = window.confirm(
      "Are you sure you want to delete this report permanently?"
    );
    if (!ok) return;

    try {
      const { error } = await supabase
        .from(table)
        .delete()
        .eq("id", editingReportId);
      if (error) throw error;

      setEditingReportId(null);
      setEditingReportCreatedBy(null);
      setShowBeforeSaleModal(false);
      setShowAfterSaleModal(false);
      setShowLayerModal(false);

      if (selectedHistoryVisit) {
        await handleHistoryClick(selectedHistoryVisit);
      }
      await loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to delete report.");
    }
  }

  // ================= PART 2A OF 2 (Paste right below Part 1) =================

  async function handleBeforeSaleSubmit() {
    if (savingBefore) return;
    setSavingBefore(true);
    try {
    if (!activeReportVisit) return;
    if (!validateBeforeSaleInputs(beforeForm)) return;
    // Permission guard (rule 5): Poultry report add needs the category
    // assigned; edits also need ownership (old rows: first edit stamps).
    if (editingReportId !== null && editingReportId !== undefined) {
      if (!canEditPoultryReport({ created_by: editingReportCreatedBy })) {
        alert("You can only edit reports you added.");
        return;
      }
    } else if (!canAddPoultryReport) {
      alert("Your assigned categories do not include Poultry, so you cannot add reports here.");
      return;
    }

    // Report validation (before any DB write): date + mortality cap.
    // Before/After reports never change t_birds stock, so the cap is the
    // batch total itself on both add and edit.
    const beforeDateErr = validateReportDate(
      reportVisitDate,
      activeReportVisit.hatch_date
    );
    if (beforeDateErr) {
      alert(beforeDateErr);
      return;
    }
    const beforeMortErr = validateMortality(
      Number(beforeForm.mortality_pcs) || 0,
      Number(activeReportVisit.t_birds) || 0
    );
    if (beforeMortErr) {
      alert(beforeMortErr);
      return;
    }

    try {
      // --- NEW CHECK FOR DUPLICATE DATE ---
      const { data: existingBefore } = await supabase
        .from("before_sale_reports")
        .select("id")
        .eq("vet_visit_id", activeReportVisit.id)
        .eq("visit_date", reportVisitDate)
        .limit(1)
        .maybeSingle();

      if (existingBefore && String(existingBefore.id) !== String(editingReportId)) {
        alert(`A Before Sale report already exists on ${formatDmy(reportVisitDate)} for this batch. Please choose a different date or edit the existing report from History.`);
        return;
      }
      // ------------------------------------

      await processFeedCompanyChangeIfNeeded(
        activeReportVisit,
        beforeForm.feed_company,
        transferAgentId,
        reportVisitDate
      );

      const cleanRemarks = String(beforeForm.remarks || "").trim();
      const updateFields: Record<string, any> = {
        visit_date: reportVisitDate,
        mortality_pcs: Number(beforeForm.mortality_pcs) || 0,
        weight_gm_b: Number(beforeForm.weight_gm_b) || 0,
        feed_t_kg: Number(beforeForm.feed_t_kg) || 0,
        feed_gm_b: Number(beforeForm.feed_gm_b) || 0,
        fcr_w_out_m: Number(beforeForm.fcr_w_out_m) || 0,
        fcr_with_m: Number(beforeForm.fcr_with_m) || 0,
        uneven: beforeForm.uneven || "",
        indigestion: beforeForm.indigestion || "",
        orange_feces: beforeForm.orange_feces || "",
        watery_feces: beforeForm.watery_feces || "",
        feeder: beforeForm.feeder || "",
        drinker: beforeForm.drinker || "",
        floor_space: beforeForm.floor_space || "",
        litter_condition: beforeForm.litter_condition || "",
        ventilation: beforeForm.ventilation || "",
        feed_scattering: beforeForm.feed_scattering || "",
        feed_hardness: beforeForm.feed_hardness || "",
        feed_company: beforeForm.feed_company,
        feed_type: beforeForm.feed_type,
        pro_date: beforeForm.pro_date || null,
        batch_no: beforeForm.batch_no || "",
        remarks: cleanRemarks,
      };

      if (editingReportId !== null && editingReportId !== undefined) {
        const { data: updatedRows, error } = await supabase
          .from("before_sale_reports")
          .update(updateFields)
          .eq("id", editingReportId)
          .select();

        if (error) throw error;
        if (!updatedRows || updatedRows.length === 0) {
          throw new Error(
            "Database blocked the update (0 rows updated). Please run the RLS UPDATE policy SQL in Supabase."
          );
        }

        setAllReportsForFarmer((prev) => ({
          ...prev,
          before: prev.before.map((r) =>
            String(r.id) === String(editingReportId)
              ? { ...r, ...updatedRows[0] }
              : r
          ),
        }));

        // Ownership (rule 2): stamp created_by on first edit of an
        // old row that has none.
        if (!editingReportCreatedBy && currentUserId) {
          try {
            await supabase
              .from("before_sale_reports")
              .update({ created_by: currentUserId })
              .eq("id", editingReportId);
          } catch {
            // created_by column not added yet.
          }
        }

        alert("Before Sale Report updated!");
      } else {
        const insertPayload = {
          ...updateFields,
          vet_visit_id: activeReportVisit.id,
          farmer_id: String(activeReportVisit.farmer_id),
          agent_id: String(transferAgentId || activeReportVisit.agent_id),
          company_id: String(activeReportVisit.company_id),
        };
        const { error } = await insertWithOptionalShedId(
          "before_sale_reports",
          insertPayload,
          activeReportVisit.shed_id
        );
        if (error) throw error;
        alert("Before Sale Report saved!");
      }

      setEditingReportId(null);
      setEditingReportCreatedBy(null);
      setShowBeforeSaleModal(false);
      await loadVetVisits();
      if (showHistoryModal && selectedHistoryVisit) {
        await handleHistoryClick(
          selectedHistoryVisit,
          selectedHistoryVisit.farmers
        );
      }
    } catch (err: any) {
      alert(err.message || "Failed to save report");
    }
    } finally {
      setSavingBefore(false);
    }
  }

  async function handleAfterSaleSoldSubmit() {
    if (savingAfter) return;
    setSavingAfter(true);
    try {
    if (!activeReportVisit) return;
    if (!validateAfterSaleInputs(afterForm)) return;
    // Permission guard (rule 5): Poultry report add needs the category
    // assigned; edits also need ownership (old rows: first edit stamps).
    if (editingReportId !== null && editingReportId !== undefined) {
      if (!canEditPoultryReport({ created_by: editingReportCreatedBy })) {
        alert("You can only edit reports you added.");
        return;
      }
    } else if (!canAddPoultryReport) {
      alert("Your assigned categories do not include Poultry, so you cannot add reports here.");
      return;
    }

    // Report validation (before any DB write): date + mortality cap.
    // Before/After reports never change t_birds stock, so the cap is the
    // batch total itself on both add and edit.
    const afterDateErr = validateReportDate(
      reportVisitDate,
      activeReportVisit.hatch_date
    );
    if (afterDateErr) {
      alert(afterDateErr);
      return;
    }
    const afterMortErr = validateMortality(
      Number(afterForm.mortality) || 0,
      Number(activeReportVisit.t_birds) || 0
    );
    if (afterMortErr) {
      alert(afterMortErr);
      return;
    }

    try {
      // --- NEW CHECK FOR DUPLICATE DATE ---
      const { data: existingAfter } = await supabase
        .from("after_sale_reports")
        .select("id")
        .eq("vet_visit_id", activeReportVisit.id)
        .eq("visit_date", reportVisitDate)
        .limit(1)
        .maybeSingle();

      if (existingAfter && String(existingAfter.id) !== String(editingReportId)) {
        alert(`An After Sale report already exists on ${formatDmy(reportVisitDate)} for this batch. Please choose a different date or edit the existing report from History.`);
        return;
      }
      // ------------------------------------

      await processFeedCompanyChangeIfNeeded(
        activeReportVisit,
        afterForm.feed_mill,
        transferAgentId,
        reportVisitDate
      );

      const cleanRemarks = String(afterForm.remarks || "").trim();
      const updateFields: Record<string, any> = {
        visit_date: reportVisitDate,
        mortality: Number(afterForm.mortality) || 0,
        weight: Number(afterForm.weight) || 0,
        feed_consumption_kg: Number(afterForm.feed_consumption_kg) || 0,
        feed_gm_b: Number(afterForm.feed_gm_b) || 0,
        feed: afterForm.feed,
        feed_mill: afterForm.feed_mill,
        fcr_w_out_m: Number(afterForm.fcr_w_out_m) || 0,
        fcr_with_m: Number(afterForm.fcr_with_m) || 0,
        remarks: cleanRemarks,
      };

      if (editingReportId !== null && editingReportId !== undefined) {
        let { data: updatedRows, error } = await supabase
          .from("after_sale_reports")
          .update(updateFields)
          .eq("id", editingReportId)
          .select();

        if (error) {
          const fallbackFields = { ...updateFields };
          delete fallbackFields.feed_gm_b;
          const retry = await supabase
            .from("after_sale_reports")
            .update(fallbackFields)
            .eq("id", editingReportId)
            .select();
          if (retry.error) throw retry.error;
          updatedRows = retry.data;
        }

        if (!updatedRows || updatedRows.length === 0) {
          throw new Error(
            "Database blocked the update (0 rows updated). Please run the RLS UPDATE policy SQL in Supabase."
          );
        }

        setAllReportsForFarmer((prev) => ({
          ...prev,
          after: prev.after.map((r) =>
            String(r.id) === String(editingReportId)
              ? { ...r, ...updatedRows![0], feed_gm_b: updateFields.feed_gm_b }
              : r
          ),
        }));

        // Ownership (rule 2): stamp created_by on first edit of an
        // old row that has none.
        if (!editingReportCreatedBy && currentUserId) {
          try {
            await supabase
              .from("after_sale_reports")
              .update({ created_by: currentUserId })
              .eq("id", editingReportId);
          } catch {
            // created_by column not added yet.
          }
        }

        alert("After Sale Report updated!");
      } else {
        const insertPayload: any = {
          ...updateFields,
          vet_visit_id: activeReportVisit.id,
          farmer_id: String(activeReportVisit.farmer_id),
          agent_id: String(transferAgentId || activeReportVisit.agent_id),
          company_id: String(activeReportVisit.company_id),
        };

        const { error } = await insertWithOptionalShedId(
          "after_sale_reports",
          insertPayload,
          activeReportVisit.shed_id
        );

        if (error) {
          delete insertPayload.feed_gm_b;
          const { error: retryErr } = await insertWithOptionalShedId(
            "after_sale_reports",
            insertPayload,
            activeReportVisit.shed_id
          );
          if (retryErr) throw retryErr;
        }

        await supabase
          .from("vet_visits")
          .update({ status: "Sold", closed_date: reportVisitDate })
          .eq("id", activeReportVisit.id);

        await supabase
          .from("farmers")
          .update({
            status: "Active",
            last_sale_date: reportVisitDate,
          })
          .eq("id", activeReportVisit.farmer_id);

        alert("Batch marked as SOLD & moved to Active status!");
      }

      setEditingReportId(null);
      setEditingReportCreatedBy(null);
      setShowAfterSaleModal(false);
      await loadVetVisits();
      if (showHistoryModal && selectedHistoryVisit) {
        await handleHistoryClick(
          selectedHistoryVisit,
          selectedHistoryVisit.farmers
        );
      }
    } catch (err: any) {
      alert(err.message || "Failed to save after sale report");
    }
    } finally {
      setSavingAfter(false);
    }
  }
  
  
  async function handleLayerReportSubmit(markAsCulled = false) {
    if (savingLayer) return;
    setSavingLayer(true);
    try {
    if (!activeReportVisit) return;
    // Permission guard (rule 5): Poultry report add needs the category
    // assigned; edits also need ownership (old rows: first edit stamps).
    if (editingReportId !== null && editingReportId !== undefined) {
      if (!canEditPoultryReport({ created_by: editingReportCreatedBy })) {
        alert("You can only edit reports you added.");
        return;
      }
    } else if (!canAddPoultryReport) {
      alert("Your assigned categories do not include Poultry, so you cannot add reports here.");
      return;
    }

    // Report validation (before any DB write, incl. the feed-company
    // change below): date + mortality cap. On edit, t_birds already had
    // the old mortality subtracted, so the cap is t_birds + old mortality
    // (editingLayerOldMort, same value the save-path recompute uses).
    const layerDateErr = validateReportDate(
      reportVisitDate,
      activeReportVisit.hatch_date
    );
    if (layerDateErr) {
      alert(layerDateErr);
      return;
    }
    const layerMortCap =
      Number(activeReportVisit.t_birds || 0) +
      (editingReportId !== null && editingReportId !== undefined
        ? editingLayerOldMort
        : 0);
    const layerMortErr = validateMortality(
      Math.max(0, Number(layerForm.mortality_pcs || 0)),
      layerMortCap
    );
    if (layerMortErr) {
      alert(layerMortErr);
      return;
    }

    if (!markAsCulled) {
      if (!validateLayerInputs(layerForm)) return;
    } else {
      if (!cullReasonText.trim()) {
        alert("Please enter the Reason for culling this layer batch.");
        return;
      }
      const ok = window.confirm(
        `Are you sure you want to mark this Layer batch as CULLED?\nReason: ${cullReasonText.trim()}`
      );
      if (!ok) return;
    }

    try {
      // --- NEW CHECK FOR DUPLICATE DATE ---
      const { data: existingLayer } = await supabase
        .from("layer_reports")
        .select("id")
        .eq("vet_visit_id", activeReportVisit.id)
        .eq("visit_date", reportVisitDate)
        .limit(1)
        .maybeSingle();

      if (existingLayer && String(existingLayer.id) !== String(editingReportId)) {
        alert(`A Layer report already exists on ${formatDmy(reportVisitDate)} for this batch. Please choose a different date or edit the existing report from History.`);
        return;
      }
      // ------------------------------------

      if (layerForm.feed_company) {
        await processFeedCompanyChangeIfNeeded(
          activeReportVisit,
          layerForm.feed_company,
          transferAgentId,
          reportVisitDate
        );
      }

      const mortNum = Math.max(0, Number(layerForm.mortality_pcs || 0));
      const prevBirds = Number(activeReportVisit.t_birds || 0);
      // EDIT FIX: activeReportVisit.t_birds already had the OLD mortality
      // subtracted when that report was added, so the pre-mortality total
      // for this report is prevBirds + oldMort. Recompute the live total
      // from scratch from that base — max(0, prevBirds + oldMort - mortNum)
      // — exactly like the add path, never by applying a delta. For a new
      // report oldMort is 0 and this is identical to updatedTBirds.
      const isEditingLayerReport =
        editingReportId !== null && editingReportId !== undefined;
      const editingLayerOldRep = isEditingLayerReport
        ? allReportsForFarmer.layer.find(
            (r: any) => String(r.id) === String(editingReportId)
          )
        : null;
      let oldMortNum = editingLayerOldRep
        ? Math.max(
            0,
            Number(
              editingLayerOldRep.mortality ??
                editingLayerOldRep.mortality_pcs ??
                0
            )
          )
        : 0;
      if (isEditingLayerReport && !editingLayerOldRep) {
        const { data: oldRep } = await supabase
          .from("layer_reports")
          .select("mortality, mortality_pcs")
          .eq("id", editingReportId)
          .maybeSingle();
        oldMortNum = Math.max(
          0,
          Number(
            (oldRep as any)?.mortality ?? (oldRep as any)?.mortality_pcs ?? 0
          )
        );
      }
      const correctedTBirds = Math.max(0, prevBirds + oldMortNum - mortNum);

      const cleanRemarks = String(layerForm.remarks || "").trim();
      const finalRemarks =
        markAsCulled && cullReasonText.trim()
          ? [cleanRemarks, `[Culled Reason: ${cullReasonText.trim()}]`]
              .filter(Boolean)
              .join(" | ")
          : cleanRemarks;

      const resolvedBreed =
        layerForm.breed ||
        activeReportVisit.breed ||
        "Shaver Star Cross 579 Brown";
      const resolvedHatchery =
        layerForm.hatchery || activeReportVisit.hatchery || "";

      const hasAnyLayerMetrics =
        Number(layerForm.b_weight_g) > 0 ||
        Number(layerForm.feed_intake_g) > 0 ||
        Number(layerForm.egg_production) > 0 ||
        Number(layerForm.egg_weight) > 0 ||
        mortNum > 0 ||
        Boolean(finalRemarks);

      if (!markAsCulled || hasAnyLayerMetrics) {
        const updateFields: Record<string, any> = {
          visit_date: reportVisitDate,
          mortality: mortNum,
          b_weight_g: Number(layerForm.b_weight_g) || 0,
          feed_total_kg: Number(layerForm.feed_total_kg) || 0,
          feed_intake_g: Number(layerForm.feed_intake_g) || 0,
          egg_production: Number(layerForm.egg_production) || 0,
          egg_weight: Number(layerForm.egg_weight) || 0,
          feed_company: layerForm.feed_company,
          feed_type: layerForm.feed_type,
          pro_date: layerForm.pro_date || null,
          batch_no: layerForm.batch_no || "",
          hatchery: resolvedHatchery,
          breed: resolvedBreed,
          disease: layerForm.disease || "",
          vaccine: activeReportVisit.last_vaccine_name || "",
          remarks: finalRemarks,
        };

        if (editingReportId !== null && editingReportId !== undefined) {
          let { data: updatedRows, error } = await supabase
            .from("layer_reports")
            .update(updateFields)
            .eq("id", editingReportId)
            .select();

          if (error) {
            const fallbackFields = { ...updateFields };
            delete fallbackFields.feed_total_kg;
            delete fallbackFields.mortality;
            const retry = await supabase
              .from("layer_reports")
              .update(fallbackFields)
              .eq("id", editingReportId)
              .select();
            if (retry.error) throw retry.error;
            updatedRows = retry.data;
          }

          if (!updatedRows || updatedRows.length === 0) {
            throw new Error(
              "Database blocked the update (0 rows updated). Please run the RLS UPDATE policy SQL in Supabase."
            );
          }

          setAllReportsForFarmer((prev) => ({
            ...prev,
            layer: prev.layer.map((r) =>
              String(r.id) === String(editingReportId)
                ? {
                    ...r,
                    ...updatedRows![0],
                    mortality: mortNum,
                    feed_total_kg: updateFields.feed_total_kg,
                  }
                : r
            ),
          }));
          // Ownership (rule 2): stamp created_by on first edit of an
          // old row that has none.
          if (!editingReportCreatedBy && currentUserId) {
            try {
              await supabase
                .from("layer_reports")
                .update({ created_by: currentUserId })
                .eq("id", editingReportId);
            } catch {
              // created_by column not added yet.
            }
          }
        } else {
          const insertPayload: any = {
            ...updateFields,
            vet_visit_id: activeReportVisit.id,
            farmer_id: String(activeReportVisit.farmer_id),
            agent_id: String(transferAgentId || activeReportVisit.agent_id),
            company_id: String(activeReportVisit.company_id),
          };

          const { error } = await insertWithOptionalShedId(
            "layer_reports",
            insertPayload,
            activeReportVisit.shed_id
          );

          if (error) {
            delete insertPayload.feed_total_kg;
            delete insertPayload.mortality;
            const { error: retryErr } = await insertWithOptionalShedId(
              "layer_reports",
              insertPayload,
              activeReportVisit.shed_id
            );
            if (retryErr) throw retryErr;
          }
        }
      }

      if (
        resolvedBreed !== activeReportVisit.breed ||
        resolvedHatchery !== activeReportVisit.hatchery
      ) {
        await supabase
          .from("vet_visits")
          .update({ breed: resolvedBreed, hatchery: resolvedHatchery })
          .eq("id", activeReportVisit.id);
      }

      // Mortality changes the live total on ADD and on EDIT alike. On edit
      // the visit/shed/farmer stocks already had oldMort subtracted, so
      // recompute each from scratch from its pre-mortality base (stock +
      // oldMort - new mortality); never subtract the new mortality again.
      // Add: oldMort is 0, so this is the old add-only behaviour.
      if (
        (!isEditingLayerReport && mortNum > 0) ||
        (isEditingLayerReport && correctedTBirds !== prevBirds)
      ) {
        await supabase
          .from("vet_visits")
          .update({ t_birds: correctedTBirds })
          .eq("id", activeReportVisit.id);

        const targetFarmerObj = statusFarmers.find(
          (f) => String(f.id) === String(activeReportVisit.farmer_id)
        );
        const currentLayerStock = Number(targetFarmerObj?.layer || prevBirds);
        let newLayerStock = Math.max(
          0,
          currentLayerStock + oldMortNum - mortNum
        );
        // Multi-shed: mortality belongs to this visit's shed. Update that
        // shed's layer stock and recompute the farmer-row legacy total as
        // the sum across sheds, so old reports keep working.
        const mortShed = getShedForVisit(activeReportVisit, targetFarmerObj);
        if (mortShed && !String(mortShed.id).startsWith("legacy-")) {
          const shedLayer = Math.max(
            0,
            Number(mortShed.layer || 0) + oldMortNum - mortNum
          );
          const mortSheds = getShedsForFarmer(targetFarmerObj);
          newLayerStock = mortSheds.reduce(
            (sum: number, s: any) =>
              sum +
              (String(s.id) === String(mortShed.id)
                ? shedLayer
                : Number(s.layer || 0)),
            0
          );
          await updateFarmerShedStock(mortShed, {
            layer: shedLayer,
          });
        }

        await supabase
          .from("farmers")
          .update({ layer: newLayerStock })
          .eq("id", activeReportVisit.farmer_id);
      }

      if (markAsCulled) {
        const visitCullUpdate: Record<string, any> = {
          status: "Culled",
          closed_date: reportVisitDate,
          t_birds: correctedTBirds,
          culled_reason: cullReasonText.trim(),
        };
        const { error: cullErr } = await supabase
          .from("vet_visits")
          .update(visitCullUpdate)
          .eq("id", activeReportVisit.id);

        if (cullErr) {
          delete visitCullUpdate.culled_reason;
          await supabase
            .from("vet_visits")
            .update(visitCullUpdate)
            .eq("id", activeReportVisit.id);
        }

        await supabase
          .from("farmers")
          .update({
            status: "Active",
            last_sale_date: reportVisitDate,
          })
          .eq("id", activeReportVisit.farmer_id);

        alert("Layer Batch marked as CULLED & moved to Active status!");
      } else {
        alert(
          editingReportId ? "Layer Report updated!" : "Layer Report saved!"
        );
      }

      setEditingReportId(null);
      setEditingReportCreatedBy(null);
      setShowLayerModal(false);
      setShowCullReasonInput(false);
      setCullReasonText("");
      await loadVetVisits();
      if (showHistoryModal && selectedHistoryVisit) {
        await handleHistoryClick(
          selectedHistoryVisit,
          selectedHistoryVisit.farmers
        );
      }
    } catch (err: any) {
      alert(err.message || "Failed to save layer report");
    }
    } finally {
      setSavingLayer(false);
    }
  }

  async function handleAddVaccine(e: React.FormEvent) {
    e.preventDefault();
    if (savingAddVaccine) return;
    setSavingAddVaccine(true);
    try {
    if (!isSuperAdmin) return;
    if (!vaccineForm.age_days || !vaccineForm.vaccine_name.trim()) return;

    const { error } = await supabase.from("vaccines").insert([
      {
        category: vaccineForm.category,
        age_days: Number(vaccineForm.age_days),
        vaccine_name: vaccineForm.vaccine_name.trim(),
        vaccine_type: vaccineForm.vaccine_type.trim(),
        route: vaccineForm.route.trim(),
      },
    ]);
    if (error) {
      alert("Error adding vaccine: " + error.message);
      return;
    }
    setVaccineForm({ ...vaccineForm, age_days: "", vaccine_name: "" });
    await loadVaccinesList();
    } finally {
      setSavingAddVaccine(false);
    }
  }

  async function handleDeleteVaccine(id: number) {
    if (!isSuperAdmin) return;
    await supabase.from("vaccines").delete().eq("id", id);
    await loadVaccinesList();
  }

  function openAddFarmerUnderAgent(agentGroup: any) {
    if (!canAddFarmerHere) {
      alert("You do not have permission to add farmers.");
      return;
    }
    // A sub-agent group preselects the sub-agent (with its parent
    // agent); an agent group preselects the agent directly.
    const subId = agentGroup.isSubAgent
      ? String(agentGroup.subAgentId || "")
      : "";
    const agentObj = allAgents.find(
      (a) => String(a.id) === String(agentGroup.parentAgentId || "")
    );
    const nourishComp = companies.find((c) =>
      c.company_name?.toLowerCase().includes("nourish")
    );
    const compId = String(
      agentObj?.company_id || nourishComp?.id || companies[0]?.id || ""
    );
    const defDist =
      agentObj?.district ||
      agentObj?.district_name ||
      (geoDistricts.length > 0 ? geoDistricts[0].name : "Mymensingh");
    const defUpa =
      agentObj?.upazila ||
      agentObj?.upazila_name ||
      (!isFullAccess && assignedUpazilas.length === 1
        ? assignedUpazilas[0]
        : "");

    setEditingFarmerId(null);
    setEditingBatchVisit(null);
    setFarmerForm({
      ...EMPTY_FARMER_FORM,
      company_id: compId,
      agent_id: agentObj ? String(agentObj.id) : "",
      // Sub-agent group: Sub Agent mode with that sub-agent preselected
      // (agent_id = its parent agent). Agent group: Agent mode.
      assign_to: subId ? "sub" : "agent",
      sub_agent_id: subId,
      district: defDist,
      upazila: defUpa,
      farm_category: "Poultry",
    });
    setFormSheds([{ ...EMPTY_SHED_FORM }]);
    setShowFarmerFormModal(true);
  }

  function handleFullEditRow(row: any) {
    const f = row.farmer;
    const v = row.visit;
    if (!canEditRecord(role, f?.created_by, currentUserId)) {
      alert("You can only edit farmers you added.");
      return;
    }
    if (v && !canEditPoultryReport(v)) {
      alert("You can only edit batches you added.");
      return;
    }

    setEditingFarmerId(f.id);
    setEditingBatchVisit(v || null);

    setFarmerForm({
      company_id: String(f.company_id || ""),
      agent_id: String(f.agent_id || ""),
      assign_to: (f as any).sub_agent_id ? "sub" : "agent",
      sub_agent_id: String((f as any).sub_agent_id || ""),
      farmer_name: f.farmer_name || "",
      mobile: f.mobile || "",
      village: f.village || "",
      union_name: f.union_name || "",
      upazila: f.upazila || "",
      district:
        f.district ||
        (geoDistricts.length > 0 ? geoDistricts[0].name : "Mymensingh"),
      farm_category: f.farm_category || "Poultry",
      broiler: String(f.broiler || ""),
      layer: String(f.layer || ""),
      sonali: String(f.sonali || ""),
      duck: String(f.duck || ""),
      quail: String(f.quail || ""),
      deshi_chicken: String(f.deshi_chicken || ""),
      dairy_cow: String(f.dairy_cow || ""),
      bull: String(f.bull || ""),
      calf_heifer: String(f.calf_heifer || ""),
      pangus: String(f.pangus || ""),
      tilapia: String(f.tilapia || ""),
      koi: String(f.koi || ""),
      shing: String(f.shing || ""),
      pabda: String(f.pabda || ""),
      gulsha: String(f.gulsha || ""),
      magur: String(f.magur || ""),
      carp: String(f.carp || ""),
      golda: String(f.golda || ""),
      bagda: String(f.bagda || ""),
      tengra: String(f.tengra || ""),
      feed_type: f.feed_type || "",
      latitude:
        f.latitude !== null && f.latitude !== undefined
          ? String(f.latitude)
          : "",
      longitude:
        f.longitude !== null && f.longitude !== undefined
          ? String(f.longitude)
          : "",
      status: f.status || "Active",
      remarks: f.remarks || "",
    });

    // Multi-shed: load this farmer's sheds (first shed = the farmer-form
    // values above; extra sheds keep their own values).
    setFormSheds(
      getShedsForFarmer(f).map((s: any, sIdx: number) => {
        const base: any = { ...EMPTY_SHED_FORM, ...s };
        base.shed_name = getShedShortName(f, s, sIdx, getShedsForFarmer(f).length);
        if (sIdx === 0) {
          base.farm_category = f.farm_category || "Poultry";
          base.latitude =
            f.latitude !== null && f.latitude !== undefined
              ? String(f.latitude)
              : "";
          base.longitude =
            f.longitude !== null && f.longitude !== undefined
              ? String(f.longitude)
              : "";
          base.feed_type = f.feed_type || "";
          SHED_STOCK_FIELDS.forEach((field) => {
            base[field] = String((f as any)[field] ?? "");
          });
        }
        return base;
      })
    );

    if (v) {
      setBatchEditSubForm({
        hatch_date: v.hatch_date || "",
        hatchery: v.hatchery || "",
        breed: v.breed || "",
        grade: v.grade || "AG",
        box_no: v.box_no || "",
        t_birds: String(v.t_birds || ""),
        category: v.category || "Broiler",
        batch_label: v.batch_label || (v as any).displayBatchName || "",
      });
    }

    setShowFarmerFormModal(true);
  }

  async function handleSaveFarmerAndBatch() {
    if (savingFarmer) return;
    setSavingFarmer(true);
    try {
    // Assignment: Agent mode needs an agent, Sub Agent mode needs a
    // sub-agent (its parent agent is derived and stored in agent_id).
    const assignOk =
      farmerForm.assign_to === "sub"
        ? !!farmerForm.sub_agent_id
        : !!farmerForm.agent_id;
    if (
      !farmerForm.company_id ||
      !assignOk ||
      !farmerForm.farmer_name.trim() ||
      !farmerForm.farm_category ||
      !farmerForm.district ||
      !farmerForm.upazila
    ) {
      alert(
        farmerForm.assign_to === "sub"
          ? "Company, District, Upazila, Sub Agent, Farmer Name, and Farm Category are required."
          : "Company, District, Upazila, Agent, Farmer Name, and Farm Category are required."
      );
      return;
    }

    // Permission guards: farmer add is NOT category-gated; farmer
    // edits are creator-only; batch (vet_visit) edits also need
    // Poultry assigned.
    let editingFarmerCreator: string | null = null;
    if (editingFarmerId) {
      try {
        const { data: existingFarmer } = await supabase
          .from("farmers")
          .select("created_by")
          .eq("id", editingFarmerId)
          .maybeSingle();
        editingFarmerCreator = (existingFarmer as any)?.created_by ?? null;
      } catch {
        // created_by column not added yet — treat as old data.
      }
      if (!canEditRecord(role, editingFarmerCreator, currentUserId)) {
        alert("You can only edit farmers you added.");
        return;
      }
    } else if (!canAddFarmerHere) {
      alert("You do not have permission to add farmers.");
      return;
    }
    if (editingBatchVisit && !canEditPoultryReport(editingBatchVisit)) {
      alert("You can only edit batches you added.");
      return;
    }

    try {
      const payload: any = {
        ...farmerForm,
        company_id: Number(farmerForm.company_id),
        agent_id: Number(farmerForm.agent_id),
        sub_agent_id: farmerForm.sub_agent_id
          ? Number(farmerForm.sub_agent_id)
          : null,
        latitude:
          farmerForm.latitude === "" ? null : Number(farmerForm.latitude),
        longitude:
          farmerForm.longitude === "" ? null : Number(farmerForm.longitude),
      };

      [...POULTRY_FIELDS, ...CATTLE_FIELDS, ...FISH_FIELDS].forEach((field) => {
        payload[field] = Number((farmerForm as any)[field] || 0);
      });

      // Multi-shed: first shed = farmer form (empty shed_name, shown as the
      // farmer name); extra sheds from formSheds. Farmer-row stock stays
      // the SUM across sheds so old reports keep working.
      // Sheds are poultry-only: other categories save the farmer
      // without sheds (farmerForm already holds their stock/location).
      const isPoultryFarmer = farmerForm.farm_category === "Poultry";
      const shedsToSave = isPoultryFarmer ? syncFormShedsWithFarmerForm() : [];
      if (isPoultryFarmer) {
        [...POULTRY_FIELDS, ...CATTLE_FIELDS, ...FISH_FIELDS].forEach((field) => {
          payload[field] = shedsToSave.reduce(
            (sum: number, s: any) => sum + Number(s[field] || 0),
            0
          );
        });
      }

      if (editingFarmerId) {
        // If the sub_agent_id column does not exist yet (migration not
        // run), retry the update without it so the save never fails
        // for a missing column.
        let fErr = (
          await supabase
            .from("farmers")
            .update(payload)
            .eq("id", editingFarmerId)
        ).error;
        if (fErr && /sub_agent_id/i.test(String(fErr.message || ""))) {
          const { sub_agent_id, ...payloadNoSub } = payload;
          fErr = (
            await supabase
              .from("farmers")
              .update(payloadNoSub)
              .eq("id", editingFarmerId)
          ).error;
        }
        if (fErr) throw fErr;

        // Ownership (rule 2): stamp created_by on first edit of an
        // old row that has none.
        if (!editingFarmerCreator && currentUserId) {
          try {
            await supabase
              .from("farmers")
              .update({ created_by: currentUserId })
              .eq("id", editingFarmerId);
          } catch {
            // created_by column not added yet.
          }
        }

        if (isPoultryFarmer) {
          await saveShedsForFarmer(
            editingFarmerId,
            farmerForm.farmer_name,
            shedsToSave,
            editingFarmerCreator || currentUserId
          );
        }

        if (editingBatchVisit) {
          const visitUpd: Record<string, any> = {
            company_id: String(farmerForm.company_id),
            agent_id: String(farmerForm.agent_id),
            hatch_date: batchEditSubForm.hatch_date,
            hatchery: batchEditSubForm.hatchery,
            breed: batchEditSubForm.breed,
            grade: batchEditSubForm.category === "Broiler" ? batchEditSubForm.grade : null,
            box_no: batchEditSubForm.box_no,
            t_birds: Number(batchEditSubForm.t_birds) || 0,
            category: batchEditSubForm.category,
            batch_label: batchEditSubForm.batch_label,
          };
          const { error: vErr } = await updateWithOptionalShedId(
            "vet_visits",
            visitUpd,
            editingBatchVisit.shed_id,
            "id",
            editingBatchVisit.id
          );
          if (vErr) {
            delete visitUpd.batch_label;
            await updateWithOptionalShedId(
              "vet_visits",
              visitUpd,
              editingBatchVisit.shed_id,
              "id",
              editingBatchVisit.id
            );
          }
        }
        alert("Farmer & Batch details updated!");
      } else {
        // If the sub_agent_id column does not exist yet (migration not
        // run), retry the insert without it so the save never fails
        // for a missing column.
        let insRes = await supabase
          .from("farmers")
          .insert([payload])
          .select("id");
        if (
          insRes.error &&
          /sub_agent_id/i.test(String(insRes.error.message || ""))
        ) {
          const { sub_agent_id, ...payloadNoSub } = payload;
          insRes = await supabase
            .from("farmers")
            .insert([payloadNoSub])
            .select("id");
        }
        const { data: insData, error: insErr } = insRes;
        if (insErr) throw insErr;
        // Resolve the new farmer id so its first shed can be saved to
        // farmer_sheds (same name+mobile lookup FarmersPage uses when the
        // insert does not return the row).
        let newFarmerId: number | null = (insData as any)?.[0]?.id || null;
        if (!newFarmerId) {
          try {
            const { data: found } = await supabase
              .from("farmers")
              .select("id")
              .eq("farmer_name", farmerForm.farmer_name.trim())
              .eq("mobile", farmerForm.mobile || "")
              .order("created_at", { ascending: false })
              .limit(1);
            newFarmerId = (found as any)?.[0]?.id || null;
          } catch {
            newFarmerId = null;
          }
        }
        if (newFarmerId) {
          // Ownership: new farmers are stamped with their creator.
          if (currentUserId) {
            try {
              await supabase
                .from("farmers")
                .update({ created_by: currentUserId })
                .eq("id", newFarmerId);
            } catch {
              // created_by column not added yet.
            }
          }
          if (isPoultryFarmer) {
            await saveShedsForFarmer(
              newFarmerId,
              farmerForm.farmer_name,
              shedsToSave,
              currentUserId
            );
          }
        }
        alert("New Farmer saved!");
      }

      setShowFarmerFormModal(false);
      setEditingFarmerId(null);
      setEditingBatchVisit(null);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save farmer details");
    }
    } finally {
      setSavingFarmer(false);
    }
  }

  function handleLocateMeInPicker() {
    if (!navigator.geolocation) {
      alert("Geolocation is not supported by your browser.");
      return;
    }
    setLocatingInPicker(true);
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        const lat = pos.coords.latitude;
        const lng = pos.coords.longitude;
        setPickerCenterOverride([lat, lng]);
        setFarmerForm((prev) => ({
          ...prev,
          latitude: lat.toFixed(6),
          longitude: lng.toFixed(6),
        }));
        setLocatingInPicker(false);
      },
      () => {
        setLocatingInPicker(false);
        alert("Unable to retrieve your location. Please allow GPS access.");
      },
      { enableHighAccuracy: true }
    );
  }

  async function handleRemoveFromRunning(row: any) {
    if (row.kind !== "batch" || !row.visit) return;
    const visit = row.visit;
    const farmer = row.farmer;
    // Removing a running batch deletes the visit + its reports.
    if (!canEditPoultryReport(visit)) {
      alert("You can only remove batches you added.");
      return;
    }

    const ok = window.confirm(
      `Warning: Remove "${
        visit.displayBatchName || farmer.farmer_name
      }" from Running?\n\nAny visit reports saved under this running batch will also be removed, and the farmer will return to Active/Inactive status based on previous history.`
    );
    if (!ok) return;

    try {
      const vId = String(visit.id);
      await Promise.all([
        supabase.from("before_sale_reports").delete().eq("vet_visit_id", vId),
        supabase.from("after_sale_reports").delete().eq("vet_visit_id", vId),
        supabase.from("layer_reports").delete().eq("vet_visit_id", vId),
      ]);

      const { error: delErr } = await supabase
        .from("vet_visits")
        .delete()
        .eq("id", visit.id);
      if (delErr) throw delErr;

      const remainingRunning = (farmer.runningVisits || []).filter(
        (rv: any) => String(rv.id) !== vId
      );

      if (remainingRunning.length === 0) {
        // Recompute from the REMAINING closed visits only — never from
        // farmers.last_sale_date (stale copy). Write the true value back
        // so the farmer row heals instead of showing a deleted sale.
        const lastSaleDate =
          farmer.latestClosed?.closed_date ||
          (farmer.latestClosed?.created_at
            ? String(farmer.latestClosed.created_at).slice(0, 10)
            : null);
        const daysSince = getDaysSince(lastSaleDate);
        const revertedStatus =
          daysSince !== null && daysSince <= 60 ? "Active" : "Inactive";

        await supabase
          .from("farmers")
          .update({ status: revertedStatus, last_sale_date: lastSaleDate })
          .eq("id", farmer.id);
      }

      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to remove batch from running.");
    }
  }

  // History Click: Only includes batches that actually have saved reports
  async function handleHistoryClick(visitOrBatch: any, farmerObj?: any) {
    const farmer =
      farmerObj ||
      statusFarmers.find(
        (f) => String(f.id) === String(visitOrBatch?.farmer_id)
      ) ||
      visitOrBatch?.farmers;

    const farmerIdStr = String(farmer?.id || visitOrBatch?.farmer_id || "");

    const [{ data: bData }, { data: aData }, { data: lData }, { data: fLogs }] =
      await Promise.all([
        supabase
          .from("before_sale_reports")
          .select("*")
          .eq("farmer_id", farmerIdStr)
          .order("visit_date", { ascending: false }),
        supabase
          .from("after_sale_reports")
          .select("*")
          .eq("farmer_id", farmerIdStr)
          .order("visit_date", { ascending: false }),
        supabase
          .from("layer_reports")
          .select("*")
          .eq("farmer_id", farmerIdStr)
          .order("visit_date", { ascending: false }),
        supabase
          .from("farmer_feed_logs")
          .select("*")
          .eq("farmer_id", farmerIdStr)
          .order("change_date", { ascending: false }),
      ]);

    const beforeList = bData || [];
    const afterList = aData || [];
    const layerList = lData || [];

    const visitIdsWithReports = new Set<string>(
      [...beforeList, ...afterList, ...layerList]
        .map((r: any) => String(r.vet_visit_id || ""))
        .filter(Boolean)
    );

    const allFarmerBatches =
      farmer?.allBatches && farmer.allBatches.length > 0
        ? farmer.allBatches
        : visitOrBatch
        ? [visitOrBatch]
        : [];

    const batchesWithReports = allFarmerBatches.filter((b: any) =>
      visitIdsWithReports.has(String(b.id))
    );

    const finalBatches =
      batchesWithReports.length > 0 ? batchesWithReports : allFarmerBatches;

    const initialBatch =
      finalBatches.find(
        (b: any) => String(b.id) === String(visitOrBatch?.id)
      ) ||
      finalBatches[0] ||
      visitOrBatch;

    setAllReportsForFarmer({
      before: beforeList,
      after: afterList,
      layer: layerList,
      feedLogs: fLogs || [],
    });
    setHistoryBatches(finalBatches);
    setSelectedHistoryVisit(initialBatch);
    setShowHistoryModal(true);
  }

  const activeBatchBeforeReports = allReportsForFarmer.before.filter(
    (r) => String(r.vet_visit_id) === String(selectedHistoryVisit?.id)
  );
  const activeBatchAfterReports = allReportsForFarmer.after.filter(
    (r) => String(r.vet_visit_id) === String(selectedHistoryVisit?.id)
  );
  const activeBatchLayerReports = allReportsForFarmer.layer.filter(
    (r) => String(r.vet_visit_id) === String(selectedHistoryVisit?.id)
  );

  const selectedBatchCulledReason = (() => {
    if (!selectedHistoryVisit) return "";
    if (selectedHistoryVisit.culled_reason) {
      return selectedHistoryVisit.culled_reason;
    }
    for (const r of activeBatchLayerReports) {
      const m = getRemarkText(r).match(/\[Culled Reason:\s*(.+?)\]/i);
      if (m && m[1]) return m[1];
    }
    return "";
  })();

  // History modal header: always show whose batch + which shed.
  // displayBatchName can be a bare stored "Batch 1" (no farmer name),
  // and the shed pill used to hide when the shed name equalled the
  // farmer name, so the header showed only "Batch 1 (Broiler • Running)".
  const histFarmerName =
    selectedHistoryVisit?.farmers?.farmer_name ||
    statusFarmers.find(
      (f: any) => String(f.id) === String(selectedHistoryVisit?.farmer_id)
    )?.farmer_name ||
    "Unknown Farmer";

  // Short batch label without the farmer-name prefix ("Raju Batch 1" ->
  // "Batch 1"; a bare "Batch 1" stays as-is).
  function getHistoryBatchShort(batch: any, farmerName: string): string {
    const raw = String(
      batch?.batch_label || batch?.displayBatchName || ""
    ).trim();
    if (raw && farmerName && farmerName !== "Unknown Farmer") {
      const lowerRaw = raw.toLowerCase();
      const lowerFarmer = farmerName.toLowerCase();
      if (lowerRaw.startsWith(lowerFarmer)) {
        const rest = raw
          .slice(farmerName.length)
          .replace(/^[\s-]+/, "")
          .trim();
        if (rest) return rest;
      }
    }
    if (raw) return raw;
    return batch?.seqNum ? `Batch ${batch.seqNum}` : "Batch";
  }

  const histBatchShort = selectedHistoryVisit
    ? getHistoryBatchShort(selectedHistoryVisit, histFarmerName)
    : "Batch";

  // Short shed label ("Raju - Shed 1" -> "Shed 1"); empty for a single
  // unnamed shed, where the farmer name itself is the shed.
  const histShedShort = (() => {
    const disp = String(selectedHistoryVisit?.shedDisplayName || "").trim();
    if (!disp || disp === histFarmerName) return "";
    const lowerDisp = disp.toLowerCase();
    const lowerFarmer = histFarmerName.toLowerCase();
    if (lowerFarmer && lowerDisp.startsWith(lowerFarmer)) {
      const rest = disp
        .slice(histFarmerName.length)
        .replace(/^[\s-]+/, "")
        .trim();
      if (rest) return rest;
    }
    return disp;
  })();

  // Native File Save & Direct Open (Tries FileOpener first so Android shows "Open with" apps instead of social Share)
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

        if (savedFile.uri) {
          const capPlugins = (Capacitor as any).Plugins;
          if (capPlugins?.FileOpener?.open) {
            await capPlugins.FileOpener.open({
              filePath: savedFile.uri,
              contentType:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              openWithDefault: true,
            });
          } else {
            await Share.share({
              url: savedFile.uri,
              dialogTitle: "Open Excel File",
            });
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
      alert("File save error: " + (err?.message || ""));
    }
  }

  // Helper to apply clean professional styling to any ExcelJS worksheet
  function styleExcelWorksheet(sheet: ExcelJS.Worksheet, headerArgb = "1E3A8A") {
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    const headerRow = sheet.getRow(1);
    headerRow.height = 24;
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: headerArgb },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
      cell.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      cell.border = {
        top: { style: "thin", color: { argb: "CBD5E1" } },
        left: { style: "thin", color: { argb: "CBD5E1" } },
        bottom: { style: "medium", color: { argb: "0F172A" } },
        right: { style: "thin", color: { argb: "CBD5E1" } },
      };
    });

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      row.height = 19;
      const bg = rowNumber % 2 === 0 ? "FFFFFF" : "F8FAFC";
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: bg },
        };
        cell.font = { color: { argb: "0F172A" }, size: 10 };
        cell.alignment = { vertical: "middle", wrapText: true };
        cell.border = {
          top: { style: "thin", color: { argb: "E2E8F0" } },
          left: { style: "thin", color: { argb: "E2E8F0" } },
          bottom: { style: "thin", color: { argb: "E2E8F0" } },
          right: { style: "thin", color: { argb: "E2E8F0" } },
        };
      });
    });

    sheet.columns.forEach((col) => {
      let maxLen = 10;
      col.eachCell?.({ includeEmpty: true }, (c) => {
        const len = String(c.value ?? "").length;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(Math.max(maxLen + 3, 11), 32);
    });
  }

  async function exportVisitHistoryExcel() {
    if (!selectedHistoryVisit) return;
    const historyShedName =
      selectedHistoryVisit.shedDisplayName ||
      selectedHistoryVisit.farmers?.farmer_name ||
      "-";
    const workbook = new ExcelJS.Workbook();
    const isLayer =
      (selectedHistoryVisit.category || "").toLowerCase() === "layer";

    if (isLayer) {
      const s = workbook.addWorksheet("Layer History");
      const headers = [
        "Shed",
        "Visit Date",
        "Hatchery",
        "H.Date",
        "Breed",
        "T.Birds",
        "Mortality",
        "Age (Wks)",
        "Wt (g)",
        "Std Wt (g)",
        "Total Feed (Kg)",
        "Feed (g/b/d)",
        "Std Feed (g)",
        "Egg%",
        "Std Egg%",
        "Egg Wt (g)",
        "Std Egg Wt",
        "Feed Company",
        "Feed Type",
        "MFG Date",
        "Batch No",
        "Disease",
        "Vaccine",
        "Remarks",
      ];
      s.addRow(headers);

      activeBatchLayerReports.forEach((r: any) => {
        const wks = getAgeWeeks(selectedHistoryVisit.hatch_date, r.visit_date);
        const breedVal =
          selectedHistoryVisit.breed || r.breed || "Shaver Star Cross 579 Brown";
        const hatcheryVal = selectedHistoryVisit.hatchery || r.hatchery || "-";
        const std = getLayerStdForVisit(breedVal, wks);

        s.addRow([
          historyShedName,
          r.visit_date || "-",
          hatcheryVal,
          selectedHistoryVisit.hatch_date || "-",
          breedVal,
          selectedHistoryVisit.t_birds || 0,
          r.mortality ?? 0,
          wks,
          r.b_weight_g || 0,
          std?.weight || "-",
          r.feed_total_kg || "-",
          r.feed_intake_g || 0,
          std?.feed_bird_day || "-",
          r.egg_production || 0,
          std?.egg_production || "-",
          r.egg_weight || 0,
          std?.egg_weight || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          r.pro_date || "-",
          r.batch_no || "-",
          r.disease || "-",
          r.vaccine || "-",
          getRemarkText(r) || "-",
        ]);
      });

      styleExcelWorksheet(s, "312E81");
    } else {
      const b = workbook.addWorksheet("Before Sale");
      b.addRow([
        "Shed",
        "Visit Date",
        "Hatchery",
        "H.Date",
        "Breed",
        "T.Birds",
        "Age (Days)",
        "Mort.",
        "Wt (g)",
        "Std Wt (g)",
        "Feed (Kg)",
        "Feed/B (g)",
        "Std Feed (g)",
        "FCR(w/o)",
        "Std FCR",
        "FCR(With)",
        "Feed Company",
        "Feed Type",
        "MFG Date",
        "Batch No",
        "Uneven%",
        "Indigestion%",
        "Orange Faeces%",
        "Watery Faeces%",
        "Litter",
        "Remarks",
      ]);

      activeBatchBeforeReports.forEach((r: any) => {
        const d = getAgeDays(selectedHistoryVisit.hatch_date, r.visit_date);
        const std = getMeatStdForVisit(selectedHistoryVisit.category, d);
        b.addRow([
          historyShedName,
          r.visit_date || "-",
          selectedHistoryVisit.hatchery || "-",
          selectedHistoryVisit.hatch_date || "-",
          selectedHistoryVisit.breed || "-",
          selectedHistoryVisit.t_birds || 0,
          d,
          r.mortality_pcs ?? 0,
          r.weight_gm_b || 0,
          std?.weight || "-",
          r.feed_t_kg || 0,
          r.feed_gm_b || 0,
          std?.feed_intake || "-",
          r.fcr_w_out_m || "-",
          std?.fcr || "-",
          r.fcr_with_m || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          r.pro_date || "-",
          r.batch_no || "-",
          r.uneven || "-",
          r.indigestion || "-",
          r.orange_feces || "-",
          r.watery_feces || "-",
          r.litter_condition || "-",
          getRemarkText(r) || "-",
        ]);
      });
      styleExcelWorksheet(b, "065F46");

      const a = workbook.addWorksheet("After Sale");
      a.addRow([
        "Shed",
        "Sale Date",
        "Hatchery",
        "H.Date",
        "Breed",
        "T.Birds",
        "Age (Days)",
        "Mort.",
        "Wt (g)",
        "Std Wt (g)",
        "Feed (Kg)",
        "Feed/B (g)",
        "Std Feed (g)",
        "FCR(w/o)",
        "Std FCR",
        "FCR(With)",
        "Feed Company",
        "Last Feed",
        "Remarks",
      ]);

      activeBatchAfterReports.forEach((r: any) => {
        const d = getAgeDays(selectedHistoryVisit.hatch_date, r.visit_date);
        const std = getMeatStdForVisit(selectedHistoryVisit.category, d);
        const tB = Number(selectedHistoryVisit.t_birds || 0);
        const fGm =
          r.feed_gm_b ||
          (tB > 0
            ? ((Number(r.feed_consumption_kg || 0) * 1000) / tB).toFixed(1)
            : "-");

        a.addRow([
          historyShedName,
          r.visit_date || "-",
          selectedHistoryVisit.hatchery || "-",
          selectedHistoryVisit.hatch_date || "-",
          selectedHistoryVisit.breed || "-",
          tB,
          d,
          r.mortality ?? 0,
          r.weight || 0,
          std?.weight || "-",
          r.feed_consumption_kg || 0,
          fGm,
          std?.feed_intake || "-",
          r.fcr_w_out_m || "-",
          std?.fcr || "-",
          r.fcr_with_m || "-",
          r.feed_mill || "-",
          r.feed || "-",
          getRemarkText(r) || "-",
        ]);
      });
      styleExcelWorksheet(a, "1E3A8A");
    }

    const batchName =
      selectedHistoryVisit.displayBatchName ||
      selectedHistoryVisit.batch_label ||
      selectedHistoryVisit.farmers?.farmer_name ||
      "Report";
    const fileName = `Batch_History_${batchName.replace(
      /[\\/:*?"<>|]/g,
      "_"
    )}.xlsx`;

    try {
      const excelBuffer = await workbook.xlsx.writeBuffer();
      await saveBufferOrBlob(excelBuffer, fileName);
    } catch (err: any) {
      console.error("Visit history export error:", err);
      alert("Visit history export failed.");
    }
  }

  const filterUpazilaOptions = Array.from(
    new Set(
      (!isFullAccess && assignedUpazilas.length > 0
        ? assignedUpazilas
        : [
            ...statusFarmers.map((f) => f.upazila),
            ...allAgents.map((a) => a.upazila),
          ]
      ).filter(Boolean)
    )
  ).sort();

  const availableAgentsForDropdown = allAgents.filter((a) => {
    if (
      selectedCompany !== "all" &&
      !getAgentCompanyIds(a).includes(String(selectedCompany))
    ) {
      return false;
    }
    if (!isFullAccess && assignedUpazilas.length > 0) {
      if (a.upazila && !assignedUpazilas.includes(a.upazila)) return false;
    }
    if (selectedUpazila !== "all") {
      const matchAgentUpa = a.upazila === selectedUpazila;
      const hasFarmerInUpa = statusFarmers.some(
        (f) =>
          String(f.agent_id) === String(a.id) && f.upazila === selectedUpazila
      );
      if (!matchAgentUpa && !hasFarmerInUpa) return false;
    }
    return true;
  });

  const modalAvailableAgents = allAgents.filter((a) => {
    if (
      farmerForm.company_id &&
      !getAgentCompanyIds(a).includes(String(farmerForm.company_id))
    ) {
      return false;
    }
    if (
      farmerForm.upazila &&
      a.upazila &&
      a.upazila !== farmerForm.upazila
    ) {
      return false;
    }
    return true;
  });

  const companyScopedFarmers = statusFarmers.filter((f) => {
    const matchComp =
      selectedCompany === "all" ||
      String(f.company_id) === String(selectedCompany);
    const matchUpa =
      selectedUpazila === "all" ||
      f.upazila === selectedUpazila ||
      f.agentUpazila === selectedUpazila;
    const matchAgent =
      selectedAgent === "all" || String(f.agent_id) === String(selectedAgent);
    // Sub-agent filter is independent of the agent filter; a farmer
    // under a sub-agent still carries the parent agent_id, so Agent +
    // Sub Agent combine as AND.
    const matchSubAgent =
      selectedSubAgent === "all" ||
      String((f as any).sub_agent_id || "") === String(selectedSubAgent);
    return matchComp && matchUpa && matchAgent && matchSubAgent;
  });

  function getSummaryForGroup(groupFilter: (f: any) => boolean) {
    const list = companyScopedFarmers.filter(groupFilter);
    const byCat: Record<string, number> = {};
    CATEGORIES.forEach((c) => {
      byCat[c] = list.filter(
        (item) => item.category?.toLowerCase() === c.toLowerCase()
      ).length;
    });
    return { total: list.length, byCat };
  }

  // ---- Per-shed table rows (one row per shed, not per farmer) ----
  // A shed with a running batch shows one row per running visit
  // (status Running, as before). Every other shed shows one row with
  // its own Active/Inactive status from its own visits, so Active and
  // Inactive sheds are visible in the table — not only Running ones.
  // A row's status is therefore always that SHED's status.
  const tableRows: any[] = [];
  companyScopedFarmers.forEach((f) => {
    const rows: any[] = [];
    (f.runningVisits || []).forEach((v: any) => {
      rows.push({
        kind: "batch",
        key: "v" + v.id,
        status: "Running",
        category: v.category || "Broiler",
        farmer: f,
        visit: v,
        shed: getShedForVisit(v, f),
        shedDisplayName: getVisitShedDisplayName(v, f),
        statusDate: v.hatch_date || "",
        daysSinceSale: f.daysSinceSale,
        closedType: "",
        isFirstRowOfFarmer: false,
        isLastBatchOfFarmer: false,
      });
    });
    (f.shedStatuses || []).forEach((ss: any) => {
      if (ss.status === "Running") return; // covered by a batch row above
      const shedList = getShedsForFarmer(f);
      rows.push({
        kind: "shed",
        key: "s" + String(ss.shed?.id ?? `${f.id}-${rows.length}`),
        status: ss.status,
        category: f.category,
        farmer: f,
        visit: null,
        shed: ss.shed,
        shedDisplayName: getShedDisplayName(
          f,
          ss.shed,
          shedList.indexOf(ss.shed),
          shedList.length
        ),
        statusDate: ss.statusDate,
        daysSinceSale: ss.daysSinceSale,
        closedType: ss.closedType,
        isFirstRowOfFarmer: false,
        isLastBatchOfFarmer: false,
      });
    });
    // Fallback: a farmer with no shed rows at all (cannot normally
    // happen — getShedsForFarmer always yields at least one) keeps the
    // old single farmer row so the farmer never disappears.
    if (rows.length === 0) {
      rows.push({
        kind: "farmer",
        key: "f" + f.id,
        status: f.computedStatus,
        category: f.category,
        farmer: f,
        visit: null,
        shed: null,
        shedDisplayName: f.farmer_name,
        statusDate: f.statusDate,
        daysSinceSale: f.daysSinceSale,
        closedType: f.closedType,
        isFirstRowOfFarmer: false,
        isLastBatchOfFarmer: false,
      });
    }
    rows[0].isFirstRowOfFarmer = true;
    rows[rows.length - 1].isLastBatchOfFarmer = true;
    rows.forEach((r: any) => tableRows.push(r));
  });

  // Summary counts are per-shed ROWS now (Total/Running/Active/
  // Inactive), matching the table. Leave/Convert stay farmer-level
  // (nourishTag is a farmer attribute) via getSummaryForGroup.
  function getRowSummaryForGroup(groupFilter: (r: any) => boolean) {
    const list = tableRows.filter(groupFilter);
    const byCat: Record<string, number> = {};
    CATEGORIES.forEach((c) => {
      byCat[c] = list.filter(
        (item) => item.category?.toLowerCase() === c.toLowerCase()
      ).length;
    });
    return { total: list.length, byCat };
  }

  const summaryCards = [
    {
      key: "all" as StatusTab,
      title: "Total",
      color: "border-slate-300 bg-white text-slate-800",
      badgeBg: "bg-slate-800 text-white",
      data: getRowSummaryForGroup(() => true),
    },
    {
      key: "Running" as StatusTab,
      title: "Running",
      color: "border-emerald-300 bg-emerald-50/60 text-emerald-900",
      badgeBg: "bg-emerald-600 text-white",
      data: getRowSummaryForGroup((r) => r.status === "Running"),
    },
    {
      key: "Active" as StatusTab,
      title: "Active (<2M)",
      color: "border-blue-300 bg-blue-50/60 text-blue-900",
      badgeBg: "bg-blue-600 text-white",
      data: getRowSummaryForGroup((r) => r.status === "Active"),
    },
    {
      key: "Inactive" as StatusTab,
      title: "Inactive (>2M)",
      color: "border-amber-300 bg-amber-50/60 text-amber-900",
      badgeBg: "bg-amber-600 text-white",
      data: getRowSummaryForGroup((r) => r.status === "Inactive"),
    },
    {
      key: "Leave from Nourish" as StatusTab,
      title: "Leave Nourish",
      color: "border-red-300 bg-red-50/60 text-red-900",
      badgeBg: "bg-red-600 text-white",
      data: getSummaryForGroup((f) => f.nourishTag === "Leave from Nourish"),
    },
    {
      key: "Convert to Nourish" as StatusTab,
      title: "Convert Nourish",
      color: "border-teal-300 bg-teal-50/60 text-teal-900",
      badgeBg: "bg-teal-600 text-white",
      data: getSummaryForGroup((f) => f.nourishTag === "Convert to Nourish"),
    },
  ];


  function rowRank(r: any) {
    const base = STATUS_RANK[r.status] ?? 3;
    if (r.status !== "Running" && r.farmer.nourishTag) return 3;
    return base;
  }

  const tableFilteredRows = tableRows
    .filter((r) => {
      if (
        selectedCategoryFilter !== "all" &&
        r.category?.toLowerCase() !== selectedCategoryFilter.toLowerCase()
      )
        return false;
      if (selectedStatusFilter !== "all") {
        if (
          selectedStatusFilter === "Leave from Nourish" ||
          selectedStatusFilter === "Convert to Nourish"
        ) {
          if (r.farmer.nourishTag !== selectedStatusFilter) return false;
        } else if (r.status !== selectedStatusFilter) return false;
      }
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const f = r.farmer;
        const match =
          f.farmer_name?.toLowerCase().includes(q) ||
          f.mobile?.toLowerCase().includes(q) ||
          f.village?.toLowerCase().includes(q) ||
          f.agentName?.toLowerCase().includes(q) ||
          String(r.visit?.batch_label || "").toLowerCase().includes(q) ||
          String(r.visit?.shedDisplayName || "")
            .toLowerCase()
            .includes(q) ||
          getShedsForFarmer(f).some((s: any, si: number) =>
            getShedDisplayName(f, s, si, getShedsForFarmer(f).length)
              .toLowerCase()
              .includes(q) ||
            getShedShortName(f, s, si, getShedsForFarmer(f).length)
              .toLowerCase()
              .includes(q)
          );
        if (!match) return false;
      }
      return true;
    })
    .sort(
      (a, b) =>
        rowRank(a) - rowRank(b) ||
        String(a.farmer.farmer_name || "").localeCompare(
          String(b.farmer.farmer_name || "")
        ) ||
        Number(a.visit?.seqNum || 0) - Number(b.visit?.seqNum || 0)
    );

  const filteredAgentsList = Object.values(
    tableFilteredRows.reduce((acc: Record<string, any>, r: any) => {
      // A farmer under a sub-agent forms that sub-agent's own group
      // (key "sub-<id>"); a farmer directly under an agent stays in
      // the agent's group (key "agent-<id>"). The Sl serial restarts
      // per group, so it now restarts per sub-agent group too.
      const subId = r.farmer.sub_agent_id
        ? String(r.farmer.sub_agent_id)
        : "";
      const key = subId
        ? `sub-${subId}`
        : `agent-${String(r.farmer.agent_id || "unknown")}`;
      if (!acc[key]) {
        const subName = subId
          ? subAgentById[subId]?.sub_agent_name || ""
          : "";
        acc[key] = {
          agentId: key,
          agentName: subId ? subName || "Sub Agent" : r.farmer.agentName,
          parentAgentName: subId ? r.farmer.agentName || "" : "",
          parentAgentId: String(r.farmer.agent_id || ""),
          isSubAgent: !!subId,
          subAgentId: subId,
          rows: [],
          counts: {} as Record<string, number>,
        };
      }
      acc[key].rows.push(r);
      acc[key].counts[r.status] = (acc[key].counts[r.status] || 0) + 1;
      return acc;
    }, {})
  ).sort((x: any, y: any) =>
    String(x.agentName).localeCompare(String(y.agentName))
  ) as any[];

  // Add a shed for the activate-modal farmer, right from the Add Batch
  // modal (same naming + previous-shed prompt as addFormShed, but the
  // shed is written to farmer_sheds immediately so it appears in the
  // shed select at once — and on FarmersPage, which reads the same table).
  // Shed add follows the farmer rule: any officer may add (no category
  // gate — only reports are category-gated).
  async function addActivateShed() {
    if (!activateFarmer || addingActivateShed) return;
    if (!canAddFarmerHere) return;
    const farmer = activateFarmer;
    const sheds = getShedsForFarmer(farmer);
    setAddingActivateShed(true);
    try {
      // Insert a shed row; stamp the creator when the column exists,
      // retry without it when it does not (same as insertWithOptionalShedId).
      const insertShedRow = async (row: any) => {
        let res: any = await supabase
          .from("farmer_sheds")
          .insert([{ ...row, created_by: currentUserId || null }])
          .select()
          .single();
        if (
          res.error &&
          /created_by/i.test(String(res.error.message || ""))
        ) {
          res = await supabase
            .from("farmer_sheds")
            .insert([row])
            .select()
            .single();
        }
        return res;
      };

      // Adding a 2nd shed: if the only existing shed has no name yet,
      // ask for its name first (empty/cancel => "Shed 1").
      const prev = sheds.length > 0 ? sheds[0] : null;
      const prevNeedsName =
        sheds.length === 1 &&
        prev &&
        getShedShortName(
          { farmer_name: farmer.farmer_name },
          prev,
          0,
          1
        ) === "";
      let baseSheds: any[] = sheds;
      if (prevNeedsName && prev) {
        const answer = window.prompt(
          "Name the previous shed (leave empty for 'Shed 1'):",
          ""
        );
        const prevName =
          answer && answer.trim()
            ? getShedShortName(
                { farmer_name: farmer.farmer_name },
                { shed_name: answer.trim() },
                0,
                2
              )
            : "Shed 1";
        const prevIsLegacy =
          prev.id == null || String(prev.id).startsWith("legacy-");
        if (prevIsLegacy) {
          // The old single shed only exists on the farmer row — save it
          // as a real shed row first, otherwise it would vanish once a
          // 2nd shed exists.
          const prevDb = shedFormToDb(
            { ...prev, shed_name: prevName },
            Number(farmer.id),
            farmer.farmer_name,
            0,
            2
          );
          const prevRes = await insertShedRow(prevDb);
          if (prevRes.error) throw prevRes.error;
          baseSheds = [prevRes.data];
        } else {
          await supabase
            .from("farmer_sheds")
            .update({ shed_name: prevName })
            .eq("id", Number(prev.id));
          baseSheds = [{ ...prev, shed_name: prevName }];
        }
      }

      // The new shed: ask for its name first — a typed name is saved,
      // empty/cancel falls back to the default "Shed N" label.
      const newCount = baseSheds.length + 1;
      const newAnswer = window.prompt(
        `Enter the new shed name (leave empty for 'Shed ${newCount}'):`,
        ""
      );
      const newName =
        newAnswer && newAnswer.trim()
          ? getShedShortName(
              { farmer_name: farmer.farmer_name },
              { shed_name: newAnswer.trim() },
              newCount - 1,
              newCount
            )
          : `Shed ${newCount}`;
      const cat = activateForm.category;
      const farmCat = cat === "Cattle" || cat === "Fish" ? cat : "Poultry";
      const newDb = shedFormToDb(
        {
          ...EMPTY_SHED_FORM,
          shed_name: newName,
          farm_category: farmCat,
          latitude: farmer.latitude != null ? String(farmer.latitude) : "",
          longitude:
            farmer.longitude != null ? String(farmer.longitude) : "",
          status: "Active",
        },
        Number(farmer.id),
        farmer.farmer_name,
        newCount - 1,
        newCount
      );
      const newRes = await insertShedRow(newDb);
      if (newRes.error) throw newRes.error;
      const newShed = newRes.data;

      // Refresh the modal select + the table behind it from local state
      // (no full reload — the modal stays open on the new shed).
      const allSheds = [...baseSheds, newShed];
      setActivateFarmer({ ...farmer, farmer_sheds: allSheds });
      setStatusFarmers((prevList: any[]) =>
        prevList.map((f: any) => {
          if (String(f.id) !== String(farmer.id)) return f;
          const updatedStatuses = (f.shedStatuses || []).map((ss: any) => {
            if (
              prevNeedsName &&
              prev &&
              String(ss.shed?.id) === String(prev.id)
            ) {
              return { ...ss, shed: baseSheds[0] };
            }
            const match = allSheds.find(
              (s: any) => String(s.id) === String(ss.shed?.id)
            );
            return match ? { ...ss, shed: match } : ss;
          });
          if (
            !updatedStatuses.some(
              (ss: any) => String(ss.shed?.id) === String(newShed.id)
            )
          ) {
            updatedStatuses.push({
              shed: newShed,
              status: "Active",
              statusDate: "",
              daysSinceSale: null,
              closedType: "",
            });
          }
          return {
            ...f,
            farmer_sheds: allSheds,
            shedStatuses: updatedStatuses,
          };
        })
      );
      setActivateShedId(String(newShed.id));
    } catch (e: any) {
      console.error("Could not add shed:", e);
      alert("Could not add shed. " + (e?.message || ""));
    } finally {
      setAddingActivateShed(false);
    }
  }

  function openActivateModal(farmer: any) {
    if (!canAddPoultryReport) {
      alert("Your assigned categories do not include Poultry, so you cannot activate batches here.");
      return;
    }
    const defaultCat = farmer.category || "Broiler";
    const defaultBreeds = BREEDS_BY_CATEGORY[defaultCat] || [];
    const defaultHatchery =
      companies.find((c) => String(c.id) === String(farmer.company_id))
        ?.company_name ||
      companies[0]?.company_name ||
      "";
    setActivateFarmer(farmer);
    setActivateShedId("");
    setActivateMode("add");
    setActivateForm({
      category: defaultCat,
      hatch_date: new Date().toISOString().split("T")[0],
      hatchery: defaultHatchery,
      breed: defaultBreeds[0] || "",
      grade: "AG",
      box_no: "",
      t_birds: "",
      batch_mark: "",
    });
    setShowActivateModal(true);
  }

  const activatePrevCats: string[] = (() => {
    if (!activateFarmer) return [];
    const running = (activateFarmer.runningVisits || [])
      .map((v: any) => v.category)
      .filter(Boolean);
    const uniq = Array.from(new Set(running)) as string[];
    if (uniq.length > 0) return uniq;
    return activateFarmer.category ? [activateFarmer.category] : [];
  })();
  const activateNeedsChoice =
    activatePrevCats.length > 0 &&
    !activatePrevCats.some(
      (c) => c.toLowerCase() === activateForm.category.toLowerCase()
    );
  const activateSheds: any[] = activateFarmer
    ? getShedsForFarmer(activateFarmer)
    : [];
  // Next batch number for the activate modal title — counts only the
  // selected shed's batches (batch numbers are per shed).
  const activateNextBatchNo =
    (activateFarmer?.allBatches || []).filter(
      (v: any) =>
        String(v.shed?.id || "none") ===
        String(activateShedId || activateSheds[0]?.id || "none")
    ).length + 1;

  async function handleActivateSubmit() {
    if (savingActivate) return;
    setSavingActivate(true);
    try {
    if (!activateForm.hatch_date || !activateForm.t_birds) {
      alert("Hatch Date and Total Birds are required.");
      return;
    }
    // Activating a batch creates a vet_visit (a Poultry report).
    if (!canAddPoultryReport) {
      alert("Your assigned categories do not include Poultry, so you cannot activate batches here.");
      return;
    }
    const targets = activateFarmer ? [activateFarmer] : [];
    if (targets.length === 0) return;

    try {
      for (const target of targets) {
        const running: any[] = target.runningVisits || [];
        const baseName = target.farmer_name || "Farmer";
        const cat = activateForm.category;
        const catLower = cat.toLowerCase();
        const qty = Number(activateForm.t_birds);
        const sameCat = running.filter(
          (v) => (v.category || "").toLowerCase() === catLower
        );
        const mode: "add" | "replace" =
          activateNeedsChoice ? activateMode : "add";

        const targetSheds = getShedsForFarmer(target);
        const chosenShed =
          targetSheds.find(
            (s: any) => String(s.id) === String(activateShedId)
          ) || targetSheds[0] || null;

        // Batch numbers count per shed: only this shed's batches.
        const totalExistingBatches = (target.allBatches || []).filter(
          (v: any) =>
            String(v.shed?.id || "none") ===
            String(chosenShed?.id || "none")
        ).length;
        const nextSeqNum = totalExistingBatches + 1;

        const manual = activateForm.batch_mark.trim() || "";
        const autoNumberedName = `${baseName} Batch ${nextSeqNum}`;
        const batchLabel = manual
          ? `${autoNumberedName} (${manual})`
          : autoNumberedName;

        const farmerUpdate: any = {
          status: "Running",
          farm_type: cat,
          farm_category:
            cat === "Cattle" || cat === "Fish" ? cat : "Poultry",
        };
        const stockKey = STOCK_KEY_BY_CATEGORY[cat];
        if (stockKey) {
          const cur = Number(target[stockKey] || 0);
          farmerUpdate[stockKey] = sameCat.length > 0 ? cur + qty : qty;
          if (mode === "replace") {
            POULTRY_STOCK_KEYS.forEach((k) => {
              if (k !== stockKey) farmerUpdate[k] = 0;
            });
          }
          // Multi-shed: the same stock change applies to the chosen shed's
          // farmer_sheds row (farmer-row totals above stay as legacy sums).
          if (chosenShed) {
            const shedStockUpdate: Record<string, any> = {};
            const shedCur = Number(chosenShed[stockKey] || 0);
            shedStockUpdate[stockKey] =
              sameCat.length > 0 ? shedCur + qty : qty;
            if (mode === "replace") {
              POULTRY_STOCK_KEYS.forEach((k) => {
                if (k !== stockKey) shedStockUpdate[k] = 0;
              });
            }
            await updateFarmerShedStock(chosenShed, shedStockUpdate);
          }
        }
        const selectedCompObj = companies.find(
          (c) => String(c.id) === String(target.company_id)
        );
        if (
          selectedCompObj?.company_name?.toLowerCase().includes("nourish") &&
          !target.last_feed_company
        ) {
          farmerUpdate.last_feed_company = "Nourish";
        }
        const { error: fErr } = await supabase
          .from("farmers")
          .update(farmerUpdate)
          .eq("id", target.id);
        if (fErr) {
          delete farmerUpdate.last_feed_company;
          await supabase
            .from("farmers")
            .update(farmerUpdate)
            .eq("id", target.id);
        }

        const insertPayload: any = {
          farmer_id: String(target.id),
          agent_id: String(target.agent_id),
          company_id: String(target.company_id),
          category: cat,
          hatch_date: activateForm.hatch_date,
          hatchery: activateForm.hatchery,
          breed: activateForm.breed,
          grade: cat === "Broiler" ? activateForm.grade : null,
          box_no: activateForm.box_no,
          t_birds: qty,
          status: "Activated",
          batch_label: batchLabel,
        };
        const { error: visitErr } = await insertWithOptionalShedId(
          "vet_visits",
          insertPayload,
          chosenShed?.id
        );
        if (visitErr) {
          delete insertPayload.batch_label;
          const { error: retryErr } = await insertWithOptionalShedId(
            "vet_visits",
            insertPayload,
            chosenShed?.id
          );
          if (retryErr) throw retryErr;
        }
      }
      alert("Farm is now Running!");
      setShowActivateModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to activate farmer");
    }
    } finally {
      setSavingActivate(false);
    }
  }

  async function exportFilteredFarmersExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Farmer Status");
      sheet.addRow([
        "SL",
        "Status",
        "Date",
        "Category",
        "Agent",
        "Farmer Name",
        "Mobile",
        "Address",
        "Hatchery",
        "H.Date",
        "Breed",
        "Birds",
        "Nourish Conversion",
      ]);

      tableFilteredRows.forEach((r, idx) => {
        const f = r.farmer;
        const v = r.visit;
        sheet.addRow([
          idx + 1,
          r.status,
          r.kind === "batch" ? v?.hatch_date || "-" : r.statusDate || f.statusDate || "-",
          r.category,
          f.agentName,
          v?.shedDisplayName ||
            r.shedDisplayName ||
            v?.displayBatchName ||
            f.farmer_name,
          f.mobile || "-",
          [f.village, f.upazila].filter(Boolean).join(", "),
          v?.hatchery || "-",
          v?.hatch_date || "-",
          v?.breed || "-",
          v?.t_birds || 0,
          f.nourishTag
            ? `${f.nourishTag}${f.nourishTagDate ? ` (${f.nourishTagDate})` : ""}`
            : "-",
        ]);
      });

      styleExcelWorksheet(sheet, "0F172A");

      const today = new Date().toISOString().slice(0, 10);
      const fileName = `Farmer_Status_${selectedStatusFilter || "All"}_${today}.xlsx`;

      const excelBuffer = await workbook.xlsx.writeBuffer();
      await saveBufferOrBlob(excelBuffer, fileName);
    } catch (err: any) {
      console.error("Filtered farmers export error:", err);
      alert("Farmer Status export failed.");
    }
  }

  const todayStr = new Date().toISOString().split("T")[0];

  const activeAgeDays = getAgeDays(
    activeReportVisit?.hatch_date,
    reportVisitDate
  );
  const activeAgeWeeks = getAgeWeeks(
    activeReportVisit?.hatch_date,
    reportVisitDate
  );

  const activeMeatStd = getMeatStdForVisit(
    activeReportVisit?.category,
    activeAgeDays
  );

  // Layer Standard uses visit's breed or layerForm's breed
  const activeLayerBreed =
    layerForm.breed ||
    activeReportVisit?.breed ||
    "Shaver Star Cross 579 Brown";
  const activeLayerStd = getLayerStdForVisit(activeLayerBreed, activeAgeWeeks);

  const beforeChangedComp = getFeedChangeCompanyObj(beforeForm.feed_company);
  const afterChangedComp = getFeedChangeCompanyObj(afterForm.feed_mill);
  const layerChangedComp = getFeedChangeCompanyObj(layerForm.feed_company);

  const activeFeedTypeOptions =
    UPDATED_FEED_TYPES_BY_CATEGORY[activeReportVisit?.category || "Broiler"] ||
    UPDATED_FEED_TYPES_BY_CATEGORY["Broiler"];

  // Live birds shown in the layer modal: on edit, recompute from scratch
  // (t_birds + old mortality - new mortality), exactly like the save path.
  const editingLayerOldMort =
    editingReportId !== null && editingReportId !== undefined
      ? Math.max(
          0,
          Number(
            allReportsForFarmer.layer.find(
              (r: any) => String(r.id) === String(editingReportId)
            )?.mortality ??
              allReportsForFarmer.layer.find(
                (r: any) => String(r.id) === String(editingReportId)
              )?.mortality_pcs ??
              0
          )
        )
      : 0;
  const liveLayerBirdsAfterMortality = Math.max(
    0,
    Number(activeReportVisit?.t_birds || 0) +
      editingLayerOldMort -
      Math.max(0, Number(layerForm.mortality_pcs || 0))
  );

  // Mortality caps for the report inputs (submit-time validation above is
  // the real gate; these only bound the number picker). Before/After
  // reports never change t_birds, so their cap is the batch total; Layer
  // on edit is t_birds + old mortality. Unset (undefined) when unknown.
  const reportMortalityCap = Number(activeReportVisit?.t_birds || 0);
  const layerMortalityCap = reportMortalityCap + editingLayerOldMort;

  const beforeExtraCols = [
    { key: "pro_date", label: "MFG Date" },
    { key: "batch_no", label: "Batch No" },
    { key: "uneven", label: "Uneven%" },
    { key: "indigestion", label: "Indigestion%" },
    { key: "orange_feces", label: "Orange Faeces%" },
    { key: "watery_feces", label: "Watery Faeces%" },
    { key: "litter_condition", label: "Litter" },
    { key: "feeder", label: "Feeder" },
    { key: "drinker", label: "Drinker" },
    { key: "floor_space", label: "Floor Space" },
    { key: "ventilation", label: "Ventilation" },
    { key: "feed_scattering", label: "Scattering" },
    { key: "feed_hardness", label: "Hardness" },
  ].filter((col) =>
    activeBatchBeforeReports.some((r) => hasValue(r[col.key]))
  );

  const layerExtraCols = [
    { key: "mortality", label: "Mort." },
    { key: "feed_total_kg", label: "Total Feed (Kg)" },
    { key: "pro_date", label: "MFG Date" },
    { key: "batch_no", label: "Batch No" },
    { key: "disease", label: "Disease" },
    { key: "vaccine", label: "Vaccine" },
  ].filter((col) =>
    activeBatchLayerReports.some((r) => hasValue(r[col.key]))
  );

  // History Batch Bar: Selected batch is ALWAYS shown upfront; remaining batches are in compact dropdown on the right
  const otherHistoryBatches = historyBatches.filter(
    (b) => String(b.id) !== String(selectedHistoryVisit?.id)
  );

  // ================= PART 2B OF 2 (Paste right below Part 2A) =================

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-4">
      {/* Top Bar */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-3 py-2 shadow-sm">
        <div className="flex items-center gap-2">
          <h1 className="text-sm sm:text-base font-bold text-slate-800">
            Poultry Visit
          </h1>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            onClick={exportFilteredFarmersExcel}
            className="bg-slate-800 hover:bg-slate-900 text-white px-2.5 py-1 rounded-lg text-[11px] font-semibold shadow-sm transition"
          >
            Export Excel
          </button>
          <button
            type="button"
            onClick={() => setShowVaccineModal(true)}
            className="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded-lg text-[11px] font-semibold shadow-sm transition"
          >
            Vaccine Schedule
          </button>
        </div>
      </div>

      {/* Compact status summary cards + Filter Bar */}
      <div className="bg-white p-2 sm:p-2.5 rounded-xl border border-slate-300 shadow-sm space-y-2">
        <div className="flex flex-wrap gap-1.5">
          {summaryCards.map((card) => {
            const isSelected = selectedStatusFilter === card.key;
            return (
              <div
                key={card.key}
                className={`flex items-center gap-1 rounded-lg border pl-2 pr-1 py-0.5 text-[11px] cursor-pointer transition ${
                  card.color
                } ${isSelected ? "ring-2 ring-emerald-600" : ""}`}
                onClick={() => setSelectedStatusFilter(card.key)}
              >
                <span className="font-bold">{card.title}</span>
                <span
                  className={`px-1.5 rounded-full font-extrabold ${card.badgeBg}`}
                >
                  {card.data.total}
                </span>
                <button
                  type="button"
                  title="Category details"
                  onClick={(e) => {
                    e.stopPropagation();
                    setExpandedCard(
                      expandedCard === card.key ? null : card.key
                    );
                  }}
                  className="px-1 font-bold opacity-70 hover:opacity-100"
                >
                  {expandedCard === card.key ? "▴" : "▾"}
                </button>
              </div>
            );
          })}
        </div>

        {expandedCard && (
          <div className="flex flex-wrap gap-1 text-[11px] bg-slate-50 border border-slate-200 rounded-lg p-1.5">
            {CATEGORIES.map((cat) => {
              const card = summaryCards.find((c) => c.key === expandedCard);
              return (
                <button
                  key={cat}
                  type="button"
                  onClick={() => {
                    setSelectedStatusFilter(expandedCard as StatusTab);
                    setSelectedCategoryFilter(cat);
                  }}
                  className="px-2 py-0.5 rounded border border-slate-300 bg-white hover:bg-slate-100"
                >
                  {cat}: <b>{card?.data.byCat[cat] || 0}</b>
                </button>
              );
            })}
          </div>
        )}

        <div className="flex flex-wrap items-center gap-1.5">
          <input
            type="text"
            placeholder="Search farmer, mobile..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full sm:w-44 rounded-md border border-slate-300 px-2 py-1 text-[11px] outline-none focus:border-emerald-500"
          />
          
          <SearchableSelect
  compact
  allValue="all"
  value={selectedCompany}
  onChange={(v) => {
    setSelectedCompany(v);
    setSelectedAgent("all");
  }}
  placeholder="All Companies"
  options={companies.map((comp) => ({
    value: String(comp.id),
    label: comp.company_name,
  }))}
/>

<SearchableSelect
  compact
  allValue="all"
  value={selectedUpazila}
  onChange={(v) => {
    setSelectedUpazila(v);
    setSelectedAgent("all");
  }}
  placeholder="All Upazilas"
  options={filterUpazilaOptions.map((upa) => ({ value: upa, label: upa }))}
/>

<SearchableSelect
  compact
  allValue="all"
  className="max-w-[160px]"
  value={selectedAgent}
  onChange={setSelectedAgent}
  placeholder="All Agents"
  options={availableAgentsForDropdown.map((ag) => ({
    value: String(ag.id),
    label: ag.agent_name,
  }))}
/>

<SearchableSelect
  compact
  allValue="all"
  className="max-w-[160px]"
  value={selectedSubAgent}
  onChange={setSelectedSubAgent}
  placeholder="All Sub Agents"
  options={subAgents.map((s: any) => ({
    value: String(s.id),
    label: s.sub_agent_name,
  }))}
/>

<SearchableSelect
  compact
  searchable={false}
  allValue="all"
  value={selectedCategoryFilter}
  onChange={setSelectedCategoryFilter}
  placeholder="All Categories"
  options={CATEGORIES.map((cat) => ({ value: cat, label: cat }))}
/>

          <button
            type="button"
            title="Reset filters"
            onClick={() => {
              setSelectedCompany("all");
              setSelectedUpazila(
                !isFullAccess && assignedUpazilas.length === 1
                  ? assignedUpazilas[0]
                  : "all"
              );
              setSelectedAgent("all");
              setSelectedSubAgent("all");
              setSelectedCategoryFilter("all");
              setSelectedStatusFilter("all");
              setSearchQuery("");
              setFocusVisitId(null);
            }}
            className="rounded-md bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2 py-1 text-[11px] font-semibold text-slate-700 transition"
          >
            Reset
          </button>
        </div>
      </div>

      {/* Agent-wise Table */}
      {filteredAgentsList.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-400">
          {_loading ? "Loading..." : "No farmers found for the selected filter."}
        </div>
      ) : (
        filteredAgentsList.map((agentGroup: any, idx: number) => {
          const countsText = Object.entries(agentGroup.counts)
            .map(([c, cnt]) => `${c}: ${cnt}`)
            .join(", ");

          return (
            <div
              key={idx}
              className="bg-white border border-slate-300 rounded-xl shadow-sm overflow-hidden"
            >
              {/* Agent Header */}
              <div className="bg-slate-100 px-3 py-1.5 border-b border-slate-300 flex flex-wrap justify-between items-center gap-2">
                <div className="flex items-center gap-2">
                  <h2 className="font-bold text-xs sm:text-sm text-slate-800">
                    {agentGroup.isSubAgent ? (
                      <span className="leading-tight">
                        <span className="block">
                          Sub Agent:{" "}
                          <span className="text-emerald-700">
                            {agentGroup.agentName}
                          </span>{" "}
                          <span className="text-[10px] font-semibold bg-white border border-slate-300 px-1.5 py-0.5 rounded text-slate-600">
                            sub-ag
                          </span>
                        </span>
                        <span className="block text-[10px] font-normal text-slate-500">
                          Agent: {agentGroup.parentAgentName}
                        </span>
                      </span>
                    ) : (
                      <>
                        Agent:{" "}
                        <span className="text-emerald-700">
                          {agentGroup.agentName}
                        </span>
                      </>
                    )}
                  </h2>
                  {canAddFarmerHere && (
                  <button
                    type="button"
                    onClick={() => openAddFarmerUnderAgent(agentGroup)}
                    title={`Add new Farmer under ${agentGroup.agentName}`}
                    className="inline-flex items-center gap-0.5 px-2 py-0.5 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-bold shadow-sm transition"
                  >
                    <Plus size={11} />
                    <UserPlus size={12} />
                  </button>
                  )}
                </div>
                <span className="text-[11px] bg-white border border-slate-300 text-slate-700 px-2 py-0.5 rounded font-semibold">
                  {agentGroup.rows.length} ({countsText})
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-50 text-slate-700 whitespace-nowrap">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1 text-center">
                        Sl.
                      </th>
                      <th className="border border-slate-300 px-2 py-1">
                        Status & Date
                      </th>
                      <th className="border border-slate-300 px-2 py-1 text-center font-bold text-emerald-700">
                        Age
                      </th>
                      <th className="border border-slate-300 px-2 py-1">
                        Category
                      </th>
                      <th className="border border-slate-300 px-2 py-1">
                        Farmer
                      </th>
                      <th className="border border-slate-300 px-2 py-1">
                        Mobile & Address
                      </th>
                      <th className="border border-slate-300 px-2 py-1">
                        H.Date
                      </th>
                      <th className="border border-slate-300 px-2 py-1">
                        Hatchery & Breed
                      </th>
                      <th className="border border-slate-300 px-2 py-1 text-center">
                        Birds
                      </th>
                      <th className="border border-slate-300 px-2 py-1">
                        Vaccine / Alert
                      </th>
                      <th className="border border-slate-300 px-2 py-1 text-center">
                        Take Report
                      </th>
                      <th className="border border-slate-300 px-2 py-1 text-center">
                        History
                      </th>
                      <th className="border border-slate-300 px-2 py-1 text-center">
                        Action
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {agentGroup.rows.map((row: any, rowIdx: number) => {
                      const farmer = row.farmer;
                      const visit = row.visit;
                      const isBatch = row.kind === "batch";
                      const category = row.category || "Broiler";
                      const isLayer = category.toLowerCase() === "layer";
                      const ageDays = isBatch
                        ? getAgeDays(visit.hatch_date)
                        : 0;
                      const isSellTime =
                        isBatch &&
                        ((category.toLowerCase() === "broiler" &&
                          ageDays >= 28) ||
                          (category.toLowerCase() === "sonali" &&
                            ageDays >= 40));
                      const vacInfo = isBatch
                        ? getVaccineInfoForVisit(visit)
                        : null;
                      const tag = isBatch
                        ? batchTagOf(
                            visit.displayBatchName || visit.batch_label,
                            farmer.farmer_name
                          )
                        : "";
                      const shedName = isBatch
                        ? getVisitShedDisplayName(visit, farmer)
                        : row.shedDisplayName || farmer.farmer_name || "N/A";
                      const isFocused =
                        isBatch && focusVisitId === String(visit.id);
                      const days = row.daysSinceSale ?? farmer.daysSinceSale;
                      const daysText =
                        days === null || days === undefined
                          ? ""
                          : days === 0
                          ? "today"
                          : `${days}d ago`;
                      const afterBtnRing = isFocused
                        ? "ring-4 ring-yellow-400 animate-pulse"
                        : "";

                      // History is ONLY active if reports exist for this specific batch (or farmer when not running)
                      const repMeta = isBatch
                        ? visitReportMeta[String(visit.id)]
                        : farmerReportMeta[String(farmer.id)];
                      const hasVisitHistory = Boolean(
                        repMeta && repMeta.count > 0
                      );

                      const actionOptionLabel =
                        row.status === "Running"
                          ? "Add Batch"
                          : row.status === "Active"
                          ? "Run Farm"
                          : "Activate";

                      return (
                        <tr
                          key={row.key}
                          data-farmer-id={farmer.id}
                          id={isBatch ? `visit-row-${visit.id}` : undefined}
                          className={`hover:bg-slate-50 leading-tight ${
                            isFocused ? "bg-yellow-50" : ""
                          }`}
                        >
                          {/* 1. SL. */}
                          <td className="border border-slate-300 px-2 py-1 text-center">
                            {rowIdx + 1}
                          </td>
                          {/* 2. STATUS & DATE */}
                          <td className="border border-slate-300 px-2 py-1 whitespace-nowrap">
                            <div className="flex items-center gap-1">
                              <span
                                className={`inline-block px-1.5 py-0 rounded-full text-[10px] font-bold ${
                                  row.status === "Running"
                                    ? "bg-emerald-100 text-emerald-800"
                                    : row.status === "Active"
                                    ? "bg-blue-100 text-blue-800"
                                    : "bg-amber-100 text-amber-800"
                                }`}
                              >
                                {row.status}
                              </span>
                              {farmer.nourishTag && (
                                <span
                                  className={`px-1 py-0 rounded text-[9px] font-bold ${
                                    farmer.nourishTag === "Convert to Nourish"
                                      ? "bg-teal-100 text-teal-800 border border-teal-300"
                                      : "bg-red-100 text-red-800 border border-red-300"
                                  }`}
                                >
                                  {farmer.nourishTag}
                                </span>
                              )}
                            </div>
                            {!isBatch && (
                              <div className="text-[10px] text-slate-500 truncate">
                                {row.statusDate ? (
                                  <>
                                    {row.closedType ||
                                      farmer.closedType ||
                                      "Sale"}
                                    : {formatDmy(row.statusDate)}
                                    {daysText && (
                                      <span className="ml-1 font-semibold text-slate-600">
                                        • {daysText}
                                      </span>
                                    )}
                                  </>
                                ) : (
                                  "No sale record"
                                )}
                              </div>
                            )}
                          </td>

                          {/* 2. AGE */}
                          <td className="border border-slate-300 px-2 py-1 text-center font-bold text-emerald-700 whitespace-nowrap">
                            {isBatch
                              ? calculateAge(visit.hatch_date, category)
                              : "-"}
                          </td>

                          {/* 3. CATEGORY */}
                          <td className="border border-slate-300 px-2 py-1 font-semibold whitespace-nowrap">
                            {category}
                          </td>

                          {/* 4. FARMER */}
                          <td className="border border-slate-300 px-2 py-1 font-bold text-slate-800 whitespace-nowrap">
                            <div>{shedName}</div>
                            {tag && (
                              <div className="text-[10px] font-semibold text-emerald-700">
                                {tag}
                              </div>
                            )}
                          </td>

                          {/* 5. MOBILE & ADDRESS */}
                          <td className="border border-slate-300 px-2 py-1 text-slate-600 whitespace-nowrap">
                            <div>{farmer.mobile || "-"}</div>
                            <div className="text-[10px] text-slate-500">
                              {[farmer.village, farmer.upazila]
                                .filter(Boolean)
                                .join(", ")}
                            </div>
                          </td>

                          {isBatch ? (
                            <>
                              {/* 6. H.DATE */}
                              <td className="border border-slate-300 px-2 py-1 whitespace-nowrap">
                                {visit.hatch_date || "-"}
                              </td>

                              {/* 7. HATCHERY & BREED */}
                              <td className="border border-slate-300 px-2 py-1 whitespace-nowrap">
                                <div className="font-medium text-slate-800">
                                  {visit.hatchery || "-"}
                                </div>
                                <div className="text-[10px] text-slate-500">
                                  {visit.breed || "-"}
                                  {category.toLowerCase() === "broiler" && visit.grade ? (
                                    <span className="ml-1 font-bold text-indigo-700 text-[9px] bg-indigo-50 border border-indigo-200 px-1 py-0 rounded">
                                      {visit.grade}
                                    </span>
                                  ) : null}
                                </div>
                              </td>

                              {/* 8. BIRDS */}
                              <td className="border border-slate-300 px-2 py-1 text-center font-bold text-slate-800 whitespace-nowrap">
                                {visit.t_birds || 0}
                              </td>

                              {/* 9. VACCINE / ALERT */}
                              <td className="border border-slate-300 px-2 py-1 whitespace-nowrap">
                                <div className="flex items-center justify-between gap-1.5">
                                  <div className="leading-tight">
                                    <div className="text-[10px] text-emerald-800 flex items-center gap-1">
                                      {visit.last_vaccine_name ? (
                                        <span>
                                          <b>Last:</b> {visit.last_vaccine_name}{" "}
                                          {visit.last_vaccine_date
                                            ? `(${formatDmy(
                                                visit.last_vaccine_date
                                              )})`
                                            : ""}
                                        </span>
                                      ) : (
                                        <span className="text-slate-400 italic">
                                          No vaccine yet
                                        </span>
                                      )}
                                      {isSellTime && (
                                        <span className="bg-red-100 border border-red-300 text-red-700 text-[9px] px-1.5 py-0 rounded font-bold">
                                          Time to Sale
                                        </span>
                                      )}
                                    </div>
                                    {vacInfo?.upcoming && (
                                      <div className="text-[10px] text-slate-600 bg-emerald-50 border border-emerald-200 px-1 rounded inline-block">
                                        <b>Next:</b>{" "}
                                        {vacInfo.upcoming.vaccine_name} (D
                                        {vacInfo.upcoming.age_days} •{" "}
                                        {formatDmy(vacInfo.upcomingDateStr)})
                                      </div>
                                    )}
                                  </div>
                                  {canEditPoultryReport(visit) && (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      handleOpenLogVaccineModal(visit)
                                    }
                                    title="Add Given Vaccine"
                                    className="w-5 h-5 shrink-0 rounded bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs flex items-center justify-center shadow-sm transition"
                                  >
                                    +
                                  </button>
                                  )}
                                </div>
                              </td>

                              {/* 10. TAKE REPORT */}
                              <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap">
                                {!canAddPoultryReport ? null : isLayer ? (
                                  <button
                                    type="button"
                                    id={`after-sale-btn-${visit.id}`}
                                    onClick={() =>
                                      openReportModal("layer_report", visit)
                                    }
                                    className={`bg-indigo-600 hover:bg-indigo-700 text-white px-2 py-0.5 rounded text-[10px] font-semibold ${afterBtnRing}`}
                                  >
                                    Layer Report
                                  </button>
                                ) : isSellTime ? (
                                  <div className="inline-flex items-center gap-1">
                                    <button
                                      type="button"
                                      id={`after-sale-btn-${visit.id}`}
                                      onClick={() =>
                                        openReportModal("after_sale", visit)
                                      }
                                      className={`bg-red-600 hover:bg-red-700 text-white px-2 py-0.5 rounded text-[10px] font-bold shadow-sm ${afterBtnRing}`}
                                    >
                                      After Sale
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() =>
                                        openReportModal("before_sale", visit)
                                      }
                                      className="bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 px-1.5 py-0.5 rounded text-[10px] font-semibold"
                                    >
                                      Before Sale
                                    </button>
                                  </div>
                                ) : (
                                  <div className="inline-flex items-center gap-1">
                                    <button
                                      type="button"
                                      onClick={() =>
                                        openReportModal("before_sale", visit)
                                      }
                                      className="bg-emerald-600 hover:bg-emerald-700 text-white px-2 py-0.5 rounded text-[10px] font-semibold shadow-sm"
                                    >
                                      Before Sale
                                    </button>
                                    <button
                                      type="button"
                                      id={`after-sale-btn-${visit.id}`}
                                      onClick={() =>
                                        openReportModal("after_sale", visit)
                                      }
                                      className={`bg-slate-100 hover:bg-slate-200 text-slate-600 border border-slate-300 px-1.5 py-0.5 rounded text-[10px] ${afterBtnRing}`}
                                    >
                                      After Sale
                                    </button>
                                  </div>
                                )}
                              </td>
                            </>
                          ) : (
                            <td
                              colSpan={5}
                              className="border border-slate-300 px-2 py-1 text-[11px] text-slate-400 italic whitespace-nowrap"
                            >
                              No running batch
                            </td>
                          )}

                          {/* 11. HISTORY (Only Active if reports exist inside) */}
                          <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap">
                            <button
                              type="button"
                              disabled={!hasVisitHistory}
                              onClick={() => {
                                if (!hasVisitHistory) return;
                                if (isBatch) {
                                  handleHistoryClick(visit, farmer);
                                } else if (farmer.latestClosed) {
                                  handleHistoryClick(
                                    farmer.latestClosed,
                                    farmer
                                  );
                                } else {
                                  handleHistoryClick(
                                    farmer.allBatches?.[0],
                                    farmer
                                  );
                                }
                              }}
                              className={`px-2 py-0.5 rounded text-[10px] font-semibold border transition ${
                                hasVisitHistory
                                  ? "bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100 cursor-pointer"
                                  : "bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed opacity-60"
                              }`}
                            >
                              History
                            </button>
                            {hasVisitHistory && repMeta?.lastDate && (
                              <div className="text-[9px] text-slate-500 font-medium">
                                Last: {formatDmy(repMeta.lastDate)}
                              </div>
                            )}
                          </td>

                          {/* 12. ACTION */}
                          <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap">
                            <select
                              value=""
                              onChange={(e) => {
                                const val = e.target.value;
                                e.target.value = "";
                                if (val === "edit") {
                                  handleFullEditRow(row);
                                } else if (val === "remove_running") {
                                  handleRemoveFromRunning(row);
                                } else if (val === "run") {
                                  openActivateModal(farmer);
                                }
                              }}
                              className="rounded-lg border border-slate-300 bg-white hover:bg-slate-50 px-2 py-0.5 text-[10px] font-bold text-slate-700 outline-none cursor-pointer shadow-sm"
                            >
                              <option value="">
                                {row.isLastBatchOfFarmer
                                  ? `${actionOptionLabel} ▾`
                                  : "Action ▾"}
                              </option>
                              {row.isLastBatchOfFarmer && canAddPoultryReport && (
                                <option value="run">
                                  + {actionOptionLabel}
                                </option>
                              )}
                              {canFullEditRow(row) && (
                                <option value="edit">✎ Edit</option>
                              )}
                              {isBatch && canEditPoultryReport(visit) && (
                                <option value="remove_running">
                                  ✖ Remove
                                </option>
                              )}
                            </select>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          );
        })
      )}

      {/* ==================================================
          ADD / FULL EDIT FARMER (+ BATCH) MODAL WITH MAP PICKER
      ================================================== */}
      {showFarmerFormModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[65] p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-5 w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h2 className="text-sm sm:text-base font-bold text-slate-800">
                {editingFarmerId
                  ? `Edit Farmer${editingBatchVisit ? " & Batch" : ""}`
                  : "Add Farmer"}
              </h2>
              <button onClick={() => setShowFarmerFormModal(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3 text-xs">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Company *
                </label>
                <SearchableSelect
                  value={farmerForm.company_id}
                  onChange={(v) =>
                    setFarmerForm({
                      ...farmerForm,
                      company_id: v,
                      agent_id: "",
                      sub_agent_id: "",
                    })
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
                  value={farmerForm.district}
                  onChange={(v) =>
                    setFarmerForm({
                      ...farmerForm,
                      district: v,
                      upazila: "",
                      union_name: "",
                      agent_id: "",
                      sub_agent_id: "",
                    })
                  }
                  placeholder="Select District"
                  options={geoDistricts.map((d) => ({
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
                  value={farmerForm.upazila}
                  onChange={(v) =>
                    setFarmerForm({
                      ...farmerForm,
                      upazila: v,
                      union_name: "",
                      agent_id: "",
                      sub_agent_id: "",
                    })
                  }
                  placeholder="Select Upazila"
                  options={geoUpazilas.map((u) => ({
                    value: u.name,
                    label: u.name,
                  }))}
                />
              </div>

              {/* Assign to: Agent or Sub Agent (checkboxes act like radios) */}
              <div className="sm:col-span-2">
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Assign Farmer To *
                </label>
                <div className="flex items-center gap-5">
                  <label className="flex items-center gap-1.5 text-xs font-semibold text-slate-700 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={farmerForm.assign_to === "agent"}
                      onChange={() =>
                        setFarmerForm({
                          ...farmerForm,
                          assign_to: "agent",
                          sub_agent_id: "",
                        })
                      }
                      className="w-3.5 h-3.5 accent-indigo-600"
                    />
                    Agent
                  </label>
                  <label className="flex items-center gap-1.5 text-xs font-semibold text-slate-700 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={farmerForm.assign_to === "sub"}
                      onChange={() =>
                        setFarmerForm({
                          ...farmerForm,
                          assign_to: "sub",
                          sub_agent_id: "",
                        })
                      }
                      className="w-3.5 h-3.5 accent-indigo-600"
                    />
                    Sub Agent
                  </label>
                </div>
              </div>

              {farmerForm.assign_to === "agent" ? (
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-1">
                    Agent *
                  </label>
                  <SearchableSelect
                    value={farmerForm.agent_id}
                    onChange={(v) =>
                      setFarmerForm({
                        ...farmerForm,
                        agent_id: v,
                        sub_agent_id: "",
                      })
                    }
                    placeholder="Select Agent"
                    options={modalAvailableAgents.map((a) => ({
                      value: String(a.id),
                      label:
                        a.agent_name +
                        (isAgentNonExclusive(a) ? " (non-ex.)" : ""),
                    }))}
                  />
                </div>
              ) : (
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-1">
                    Sub Agent *
                  </label>
                  <SearchableSelect
                    value={farmerForm.sub_agent_id}
                    onChange={(v) => {
                      // The farmer keeps agent_id = the sub-agent's parent.
                      const parent = subAgents.find(
                        (s: any) => String(s.id) === String(v)
                      );
                      setFarmerForm({
                        ...farmerForm,
                        sub_agent_id: v,
                        agent_id: parent
                          ? String(parent.agent_id)
                          : farmerForm.agent_id,
                      });
                    }}
                    placeholder="Select Sub Agent"
                    options={subAgents.map((s: any) => ({
                      value: String(s.id),
                      label: getSubAgentOptionLabel(s),
                    }))}
                  />
                </div>
              )}

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Farmer Name *
                </label>
                <input
                  placeholder="Farmer Name"
                  value={farmerForm.farmer_name}
                  onChange={(e) =>
                    setFarmerForm({
                      ...farmerForm,
                      farmer_name: e.target.value,
                    })
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
                  value={farmerForm.mobile}
                  onChange={(e) =>
                    setFarmerForm({ ...farmerForm, mobile: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Union
                </label>
                <SearchableSelect
                  value={farmerForm.union_name}
                  onChange={(v) => setFarmerForm({ ...farmerForm, union_name: v })}
                  placeholder="Select Union"
                  options={geoUnions.map((u) => ({
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
                  value={farmerForm.village}
                  onChange={(e) =>
                    setFarmerForm({ ...farmerForm, village: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>
            </div>

            {/* Farm Category — sheds are for poultry only */}
            <div className="mb-3">
              <label className="block text-xs font-bold text-slate-600 mb-1">
                Farm Category *
              </label>
              <SearchableSelect
                searchable={false}
                value={farmerForm.farm_category || ""}
                onChange={(v) =>
                  setFarmerForm({ ...farmerForm, farm_category: v })
                }
                placeholder="Select Category"
                options={[
                  { value: "Poultry", label: "🐔 Poultry" },
                  { value: "Cattle", label: "🐄 Cattle" },
                  { value: "Fish", label: "🐟 Fish" },
                ]}
              />
              {farmerForm.farm_category &&
                farmerForm.farm_category !== "Poultry" && (
                  <p className="mt-1 text-[11px] text-slate-500">
                    Sheds are for poultry only — this farmer will be saved
                    without sheds.
                  </p>
                )}
            </div>

            {/* Sheds — one farmer, multiple sheds, each with its own location/category/stock */}
            {farmerForm.farm_category === "Poultry" && (
              <>
            <div className="mb-3">
              <div className="flex items-center justify-between mb-2">
                <label className="block text-xs font-bold text-slate-600">
                  Sheds * ({formSheds.length})
                </label>
              </div>

              {formSheds.map((shed: any, sIdx: number) => {
                const shedCat =
                  sIdx === 0
                    ? farmerForm.farm_category || ""
                    : shed.farm_category || farmerForm.farm_category || "";
                const shedShort = getShedShortName(
                  { farmer_name: farmerForm.farmer_name } as any,
                  shed,
                  sIdx,
                  formSheds.length
                );
                const stockVal = (field: string) =>
                  sIdx === 0 ? (farmerForm as any)[field] : shed[field];
                const setStockVal = (field: string, value: string) => {
                  if (sIdx === 0) {
                    setFarmerForm({ ...farmerForm, [field]: value });
                  } else {
                    updateFormShedField(sIdx, field, value);
                  }
                };
                const feedOptions = (
                  UPDATED_FEED_TYPES_BY_CATEGORY[shedCat] ||
                  UPDATED_FEED_TYPES_BY_CATEGORY["Broiler"] ||
                  []
                ).map((ft) => ({ value: ft, label: ft }));
                return (
                  <div
                    key={shed.id || sIdx}
                    className="border border-slate-300 rounded-lg p-3 mb-2 bg-slate-50"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold text-slate-700">
                        Shed {sIdx + 1}
                        {shedShort ? ` — ${shedShort}` : ""}
                      </span>
                      <button
                        type="button"
                        onClick={() => removeFormShed(sIdx)}
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
                            updateFormShedField(
                              sIdx,
                              "shed_name",
                              e.target.value
                            )
                          }
                          className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">
                          Farm Category
                        </label>
                        <div className="w-full rounded border border-slate-200 bg-slate-100 px-2.5 py-1.5 text-xs font-semibold text-slate-700">
                          🐔 Poultry
                        </div>
                      </div>
                    </div>

                    {sIdx === 0 ? (
                      <div className="mb-2 flex flex-wrap items-center gap-2 bg-white border border-slate-200 rounded p-2.5">
                        <button
                          type="button"
                          onClick={() => {
                            setPickerCenterOverride(null);
                            setShowMapPicker(true);
                          }}
                          className="inline-flex items-center gap-1.5 bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-semibold"
                        >
                          <MapPin size={13} /> Select Location from Map
                        </button>
                        <button
                          type="button"
                          onClick={handleLocateMeInPicker}
                          disabled={locatingInPicker}
                          className="inline-flex items-center gap-1.5 bg-slate-600 text-white px-3 py-1.5 rounded text-xs font-semibold disabled:opacity-50"
                        >
                          <Navigation size={13} /> Locate Me
                        </button>
                        {farmerForm.latitude && farmerForm.longitude && (
                          <span className="text-emerald-700 text-xs font-semibold">
                            📍 Selected (
                            {Number(farmerForm.latitude).toFixed(4)}, {" "}
                            {Number(farmerForm.longitude).toFixed(4)})
                          </span>
                        )}
                      </div>
                    ) : (
                      <div className="mb-2">
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">
                          Location (lat, lng)
                        </label>
                        <div className="flex gap-1">
                          <input
                            type="text"
                            placeholder="Latitude"
                            value={shed.latitude || ""}
                            onChange={(e: any) =>
                              updateFormShedField(
                                sIdx,
                                "latitude",
                                e.target.value
                              )
                            }
                            className="w-full rounded border border-slate-300 px-2 py-1 text-xs"
                          />
                          <input
                            type="text"
                            placeholder="Longitude"
                            value={shed.longitude || ""}
                            onChange={(e: any) =>
                              updateFormShedField(
                                sIdx,
                                "longitude",
                                e.target.value
                              )
                            }
                            className="w-full rounded border border-slate-300 px-2 py-1 text-xs"
                          />
                        </div>
                      </div>
                    )}

                    {shedCat === "Poultry" && (
                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 mb-2 bg-amber-50 border border-amber-200 rounded p-2.5">
                        {POULTRY_FIELDS.map((item) => (
                          <div key={item}>
                            <label className="block text-[11px] font-semibold capitalize">
                              {item.replace("_", " ")}
                            </label>
                            <input
                              type="number"
                              min="0"
                              value={stockVal(item) || ""}
                              onChange={(e) => setStockVal(item, e.target.value)}
                              className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                            />
                          </div>
                        ))}
                      </div>
                    )}

                    {shedCat === "Cattle" && (
                      <div className="grid grid-cols-3 gap-2 mb-2 bg-emerald-50 border border-emerald-200 rounded p-2.5">
                        {CATTLE_FIELDS.map((item) => (
                          <div key={item}>
                            <label className="block text-[11px] font-semibold capitalize">
                              {item.replace("_", " ")}
                            </label>
                            <input
                              type="number"
                              min="0"
                              value={stockVal(item) || ""}
                              onChange={(e) => setStockVal(item, e.target.value)}
                              className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                            />
                          </div>
                        ))}
                      </div>
                    )}

                    {shedCat === "Fish" && (
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-2 bg-sky-50 border border-sky-200 rounded p-2.5">
                        {FISH_FIELDS.map((item) => (
                          <div key={item}>
                            <label className="block text-[11px] font-semibold capitalize">
                              {item}
                            </label>
                            <input
                              type="number"
                              min="0"
                              value={stockVal(item) || ""}
                              onChange={(e) => setStockVal(item, e.target.value)}
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
                      <SearchableSelect
                        value={
                          sIdx === 0 ? farmerForm.feed_type : shed.feed_type || ""
                        }
                        onChange={(v) =>
                          sIdx === 0
                            ? setFarmerForm({ ...farmerForm, feed_type: v })
                            : updateFormShedField(sIdx, "feed_type", v)
                        }
                        placeholder="Select Feed Type"
                        options={feedOptions}
                      />
                    </div>
                  </div>
                );
              })}

              <button
                type="button"
                onClick={addFormShed}
                className="inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold"
              >
                <Plus size={13} /> Add Shed
              </button>
            </div>
              </>
            )}

            {farmerForm.farm_category &&
              farmerForm.farm_category !== "Poultry" && (
                <div className="mb-3 border border-slate-300 rounded-lg p-3 bg-slate-50">
                  <div className="text-xs font-bold text-slate-700 mb-2">
                    Location
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      type="button"
                      onClick={() => {
                        setPickerCenterOverride(null);
                        setShowMapPicker(true);
                      }}
                      className="inline-flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded text-xs font-semibold cursor-pointer"
                    >
                      <MapPin size={13} /> Select Location from Map
                    </button>
                    <button
                      type="button"
                      onClick={handleLocateMeInPicker}
                      disabled={locatingInPicker}
                      className="inline-flex items-center gap-1.5 bg-slate-600 hover:bg-slate-700 text-white px-3 py-1.5 rounded text-xs font-semibold disabled:opacity-50 cursor-pointer"
                    >
                      <Navigation size={13} /> Locate Me
                    </button>
                    {farmerForm.latitude && farmerForm.longitude && (
                      <span className="text-[11px] font-semibold text-emerald-700">
                        📍 {farmerForm.latitude}, {farmerForm.longitude}
                      </span>
                    )}
                  </div>
                </div>
              )}

            <div className="grid grid-cols-1 gap-3 mb-3 text-xs">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Remarks
                </label>
                <input
                  placeholder="Remarks"
                  value={farmerForm.remarks}
                  onChange={(e) =>
                    setFarmerForm({ ...farmerForm, remarks: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>
            </div>
            {editingBatchVisit && (
              <div className="bg-emerald-50/70 border border-emerald-200 rounded-xl p-3 mb-3 space-y-2.5 text-xs">
                <div className="font-bold text-emerald-900 border-b border-emerald-200 pb-1">
                  Running Batch Details
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-0.5">
                      Batch Name / Mark
                    </label>
                    <input
                      type="text"
                      value={batchEditSubForm.batch_label}
                      onChange={(e) =>
                        setBatchEditSubForm({
                          ...batchEditSubForm,
                          batch_label: e.target.value,
                        })
                      }
                      className="w-full rounded border border-slate-300 bg-white px-2 py-1"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold text-slate-700 mb-0.5">
                      Batch Category
                    </label>
                    <select
                      value={batchEditSubForm.category}
                      onChange={(e) => {
                        const cat = e.target.value;
                        setBatchEditSubForm({
                          ...batchEditSubForm,
                          category: cat,
                          breed: (BREEDS_BY_CATEGORY[cat] || [])[0] || "",
                        });
                      }}
                      className="w-full rounded border border-slate-300 bg-white px-2 py-1"
                    >
                      {CATEGORIES.map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block font-semibold text-slate-700 mb-0.5">
                      Hatch Date
                    </label>
                    <input
                      type="date"
                      value={batchEditSubForm.hatch_date}
                      onChange={(e) =>
                        setBatchEditSubForm({
                          ...batchEditSubForm,
                          hatch_date: e.target.value,
                        })
                      }
                      className="w-full rounded border border-slate-300 bg-white px-2 py-1"
                    />
                  </div>
                 <div>
                    <label className="block font-semibold text-slate-700 mb-0.5">
                      Hatchery
                    </label>
                    <SearchableSelect
                      value={batchEditSubForm.hatchery}
                      onChange={(v) =>
                        setBatchEditSubForm({
                          ...batchEditSubForm,
                          hatchery: v,
                        })
                      }
                      placeholder="Select Hatchery"
                      options={companies.map((c) => ({
                        value: c.company_name,
                        label: c.company_name,
                      }))}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-0.5">
                      Breed
                    </label>
                    <SearchableSelect
                    searchable={false}
                      value={batchEditSubForm.breed}
                      onChange={(v) =>
                        setBatchEditSubForm({
                          ...batchEditSubForm,
                          breed: v,
                        })
                      }
                      placeholder="Select Breed"
                      options={(
                        BREEDS_BY_CATEGORY[batchEditSubForm.category] || []
                      ).map((b) => ({
                        value: b,
                        label: b,
                      }))}
                    />
                  </div>

                  {batchEditSubForm.category === "Broiler" && (
                    <div>
                    <label className="block font-semibold text-slate-700 mb-0.5">
                      Grade
                    </label>
                    <SearchableSelect
                    searchable={false}
                      value={batchEditSubForm.grade}
                      onChange={(v) =>
                        setBatchEditSubForm({
                          ...batchEditSubForm,
                          grade: v,
                        })
                      }
                      placeholder="Select Grade"
                      options={[
                        { value: "AG", label: "AG" },
                        { value: "B+", label: "B+" },
                      ]}
                    />
                  </div>
                  )}
                  
                  <div>
                    <label className="block font-semibold text-slate-700 mb-0.5">
                      Total Birds
                    </label>
                    <input
                      type="number"
                      value={batchEditSubForm.t_birds}
                      onChange={(e) =>
                        setBatchEditSubForm({
                          ...batchEditSubForm,
                          t_birds: e.target.value,
                        })
                      }
                      className="w-full rounded border border-slate-300 bg-white px-2 py-1"
                    />
                  </div>
                </div>
              </div>
            )}

            <div className="flex justify-end gap-2 pt-2 border-t text-xs">
              <button
                type="button"
                onClick={() => setShowFarmerFormModal(false)}
                className="px-4 py-1.5 rounded border border-slate-300 font-semibold"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSaveFarmerAndBatch}
                disabled={savingFarmer}
                className="px-4 py-1.5 rounded bg-emerald-600 hover:bg-emerald-700 text-white font-semibold cursor-pointer disabled:opacity-60"
              >
                {savingFarmer
                  ? "Saving..."
                  : editingFarmerId
                  ? "Update Farmer"
                  : "Save Farmer"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MAP LOCATION PICKER MODAL FOR ADD/EDIT FARMER
      ================================================== */}
      {showMapPicker && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[75] p-2 sm:p-4">
          <div className="bg-white p-3 sm:p-4 rounded-xl w-full max-w-4xl shadow-2xl space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="font-bold text-sm sm:text-base">
                Select Farm Location
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

                {farmerForm.latitude && farmerForm.longitude && (
                  <button
                    type="button"
                    onClick={() => setShowMapPicker(false)}
                    className="inline-flex items-center gap-1 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                  >
                    <CheckCircle2 size={13} /> Use Selected Location
                  </button>
                )}

                <button onClick={() => setShowMapPicker(false)}>
                  <X size={18} />
                </button>
              </div>
            </div>

            <LocationPicker
              key={
                pickerCenterOverride
                  ? `${pickerCenterOverride[0]}_${pickerCenterOverride[1]}`
                  : "visit_default_picker"
              }
              center={
                pickerCenterOverride
                  ? pickerCenterOverride
                  : formAgentCenter
                  ? formAgentCenter
                  : selectedFormUpazila &&
                    selectedFormUpazila.lat &&
                    selectedFormUpazila.lon
                  ? [
                      Number(selectedFormUpazila.lat),
                      Number(selectedFormUpazila.lon),
                    ]
                  : selectedFormDistrict &&
                    selectedFormDistrict.lat &&
                    selectedFormDistrict.lon
                  ? [
                      Number(selectedFormDistrict.lat),
                      Number(selectedFormDistrict.lon),
                    ]
                  : [24.7471, 90.4203]
              }
              zoom={
                pickerCenterOverride
                  ? 17
                  : formAgentCenter
                  ? 14
                  : selectedFormUpazila
                  ? 13
                  : selectedFormDistrict
                  ? 11
                  : 7
              }
              selectedPosition={
                farmerForm.latitude && farmerForm.longitude
                  ? [
                      Number(farmerForm.latitude),
                      Number(farmerForm.longitude),
                    ]
                  : null
              }
              onSelect={(lat, lng) => {
                setFarmerForm({
                  ...farmerForm,
                  latitude: lat.toString(),
                  longitude: lng.toString(),
                });
                setShowMapPicker(false);
              }}
              markers={
                <>
                  {statusFarmers.flatMap((farmer: any) => {
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

      {/* ==================================================
          ACTIVATE / RUN FARM / ADD BATCH MODAL
      ================================================== */}
      {showActivateModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[60] p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-md max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h3 className="text-sm sm:text-base font-bold text-slate-800">
                {activateFarmer?.runningCount > 0
                  ? `Add Batch (${
                      activateNextBatchNo
                    }): ${activateFarmer?.farmer_name}`
                  : activateFarmer?.computedStatus === "Active"
                  ? `Run Farm (Batch ${
                      activateNextBatchNo
                    }): ${activateFarmer?.farmer_name}`
                  : `Activate Farmer (Batch ${
                      activateNextBatchNo
                    }): ${activateFarmer?.farmer_name}`}
              </h3>
              <button
                onClick={() => setShowActivateModal(false)}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <div className="space-y-3 text-xs">
  <div>
    <label className="block font-bold text-slate-600 mb-1">
      Category / Farm Type *
    </label>
    <SearchableSelect
      searchable={false}
      value={activateForm.category}
      onChange={(cat) => {
        if (!cat) return; // category khali kora jabe na
        const breeds = BREEDS_BY_CATEGORY[cat] || [];
        setActivateForm({
          ...activateForm,
          category: cat,
          breed: breeds[0] || "",
        });
      }}
      placeholder="Select Category"
      options={CATEGORIES.map((cat) => ({ value: cat, label: cat }))}
    />
  </div>

  {activateFarmer && activateSheds.length > 1 && (
    <div>
      <label className="block font-bold text-slate-600 mb-1">
        Shed *
      </label>
      <select
        value={activateShedId || String(activateSheds[0]?.id || "")}
        onChange={(e: any) => setActivateShedId(e.target.value)}
        className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 outline-none"
      >
        {activateSheds.map((s: any, sIdx: number) => (
          <option key={String(s.id)} value={String(s.id)}>
            {getShedDisplayName(
              activateFarmer,
              s,
              sIdx,
              activateSheds.length
            )}
          </option>
        ))}
      </select>
      {canAddFarmerHere && (
        <button
          type="button"
          onClick={addActivateShed}
          disabled={addingActivateShed}
          className="mt-1 text-[11px] font-bold text-emerald-700 hover:text-emerald-800 disabled:opacity-50"
        >
          {addingActivateShed ? "Adding..." : "+ Add Shed"}
        </button>
      )}
    </div>
  )}

  {activateFarmer && activateSheds.length <= 1 && (
    <div>
      <label className="block font-bold text-slate-600 mb-1">Shed</label>
      <div className="w-full rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-slate-700 font-semibold">
        {activateSheds.length > 0
          ? getShedDisplayName(
              activateFarmer,
              activateSheds[0],
              0,
              activateSheds.length
            )
          : activateFarmer.farmer_name}
      </div>
      <p className="text-[11px] text-slate-500 mt-1">
        Single shed — used automatically for this batch.
      </p>
      {canAddFarmerHere && (
        <button
          type="button"
          onClick={addActivateShed}
          disabled={addingActivateShed}
          className="mt-1 text-[11px] font-bold text-emerald-700 hover:text-emerald-800 disabled:opacity-50"
        >
          {addingActivateShed ? "Adding..." : "+ Add Shed"}
        </button>
      )}
    </div>
  )}

  {activateNeedsChoice && (
    <div className="border border-amber-300 bg-amber-50 rounded-xl p-3 space-y-2">
      <div className="font-bold text-amber-900">
        Before: {activatePrevCats.join(", ")} → Now:{" "}
        {activateForm.category}
      </div>
      <label className="flex items-start gap-2 cursor-pointer">
        <input
          type="radio"
          name="actMode"
          checked={activateMode === "add"}
          onChange={() => setActivateMode("add")}
          className="mt-0.5 accent-amber-600"
        />
        <span>
          <b>Add</b> new {activateForm.category} batch (keep{" "}
          {activatePrevCats.join(", ")})
        </span>
      </label>
      <label className="flex items-start gap-2 cursor-pointer">
        <input
          type="radio"
          name="actMode"
          checked={activateMode === "replace"}
          onChange={() => setActivateMode("replace")}
          className="mt-0.5 accent-amber-600"
        />
        <span>
          <b>Replace</b> {activatePrevCats.join(", ")} with{" "}
          {activateForm.category} on Farmers page
        </span>
      </label>
    </div>
  )}

  <div>
    <label className="block font-bold text-slate-600 mb-1">
      Hatch Date *
    </label>
    <input
      type="date"
      value={activateForm.hatch_date}
      onChange={(e) =>
        setActivateForm({
          ...activateForm,
          hatch_date: e.target.value,
        })
      }
      className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none"
    />
  </div>

  <div>
    <label className="block font-bold text-slate-600 mb-1">
      Hatchery (All Companies) *
    </label>
    <SearchableSelect
      value={activateForm.hatchery}
      onChange={(v) => setActivateForm({ ...activateForm, hatchery: v })}
      placeholder="Select Hatchery Company"
      options={companies.map((c) => ({
        value: c.company_name,
        label: c.company_name,
      }))}
    />
  </div>

  <div>
    <label className="block font-bold text-slate-600 mb-1">
      Breed *
    </label>
    <SearchableSelect
      value={activateForm.breed}
      onChange={(v) => setActivateForm({ ...activateForm, breed: v })}
      placeholder="Select Breed"
      options={(BREEDS_BY_CATEGORY[activateForm.category] || []).map((b) => ({
        value: b,
        label: b,
      }))}
    />
  </div>

  {activateForm.category === "Broiler" && (
    <div>
      <label className="block font-bold text-slate-600 mb-1">
        Grade
      </label>
      <SearchableSelect
        searchable={false}
        value={activateForm.grade}
        onChange={(v) => {
          if (!v) return; // grade khali kora jabe na
          setActivateForm({ ...activateForm, grade: v });
        }}
        placeholder="Select Grade"
        options={[
          { value: "AG", label: "AG" },
          { value: "B+", label: "B+" },
        ]}
      />
    </div>
  )}

  <div className="grid grid-cols-2 gap-2.5">
    <div>
      <label className="block font-bold text-slate-600 mb-1">
        Box No.
      </label>
      <input
        type="text"
        placeholder="Box No"
        value={activateForm.box_no}
        onChange={(e) =>
          setActivateForm({
            ...activateForm,
            box_no: e.target.value,
          })
        }
        className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none"
      />
    </div>
    <div>
      <label className="block font-bold text-slate-600 mb-1">
        Total Birds / Qty *
      </label>
      <input
        type="number"
        placeholder="Total Birds"
        value={activateForm.t_birds}
        onChange={(e) =>
          setActivateForm({
            ...activateForm,
            t_birds: e.target.value,
          })
        }
        className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none"
      />
    </div>
  </div>

    <div>
      <label className="block font-bold text-slate-600 mb-1">
        Batch Mark (Optional — Auto:{" "}
        {activateFarmer?.farmer_name} Batch{" "}
        {activateNextBatchNo})
      </label>
      <input
        type="text"
        placeholder="Optional custom note e.g. Shed-2"
        value={activateForm.batch_mark}
        onChange={(e) =>
          setActivateForm({
            ...activateForm,
            batch_mark: e.target.value,
          })
        }
        className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none"
      />
    </div>

  <div className="flex justify-end gap-2 pt-3 border-t">
    <button
      type="button"
      onClick={() => setShowActivateModal(false)}
      className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 font-semibold"
    >
      Cancel
    </button>
    <button
      type="button"
      onClick={handleActivateSubmit}
      disabled={savingActivate}
      className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold shadow-sm cursor-pointer disabled:opacity-60"
    >
      {savingActivate ? "Activating..." : "Run Farm"}
    </button>
  </div>
</div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 0: QUICK "+" ADD GIVEN VACCINE MODAL
      ================================================== */}
      {showLogVaccineModal && activeVaccineVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-sm shadow-2xl border border-slate-300 space-y-3.5 text-xs">
            <div className="flex justify-between items-center border-b pb-2">
              <div>
                <h3 className="text-sm font-bold text-slate-800">
                  Add Given Vaccine
                </h3>
                <p className="text-[11px] text-slate-500">
                  {activeVaccineVisit.displayBatchName ||
                    activeVaccineVisit.batch_label ||
                    activeVaccineVisit.farmers?.farmer_name}{" "}
                  ({activeVaccineVisit.category})
                </p>
              </div>
              <button
                onClick={() => setShowLogVaccineModal(false)}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <form onSubmit={handleSaveLoggedVaccine} className="space-y-3">
              <div>
                <label className="block font-bold text-indigo-800 mb-1">
                  Select from Schedule (Optional)
                </label>
                <select
                  value=""
                  onChange={(e) => {
                    if (e.target.value) {
                      setLogVaccineForm({
                        ...logVaccineForm,
                        vaccine_name: e.target.value,
                      });
                    }
                  }}
                  className="w-full border border-indigo-300 rounded-lg px-2.5 py-1.5 bg-indigo-50/40"
                >
                  <option value="">
                    -- Pick scheduled vaccine or type below --
                  </option>
                  {vaccines
                    .filter(
                      (v) =>
                        v.category?.toLowerCase() ===
                        (activeVaccineVisit.category || "Broiler").toLowerCase()
                    )
                    .map((v, i) => (
                      <option
                        key={i}
                        value={`${v.vaccine_name} (${v.route})`}
                      >
                        Day {v.age_days}: {v.vaccine_name} — {v.route}
                      </option>
                    ))}
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">
                  Vaccine Name *
                </label>
                <input
                  type="text"
                  required
                  placeholder="e.g. IB + ND (Eye Drop)"
                  value={logVaccineForm.vaccine_name}
                  onChange={(e) =>
                    setLogVaccineForm({
                      ...logVaccineForm,
                      vaccine_name: e.target.value,
                    })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-indigo-600"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">
                  Vaccine Date *
                </label>
                <input
                  type="date"
                  required
                  min={activeVaccineVisit?.hatch_date || undefined}
                  max={todayStr}
                  value={logVaccineForm.vaccine_date}
                  onChange={(e) =>
                    setLogVaccineForm({
                      ...logVaccineForm,
                      vaccine_date: e.target.value,
                    })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-indigo-600"
                />
              </div>

              <div className="flex justify-end gap-2 pt-2 border-t">
                <button
                  type="button"
                  onClick={() => setShowLogVaccineModal(false)}
                  className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 font-semibold"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={savingVaccine}
                  className="px-4 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white font-bold shadow-sm cursor-pointer disabled:opacity-60"
                >
                  {savingVaccine ? "Saving..." : "OK / Add"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 1: BEFORE SALE REPORT (Supports both New & Edit from History + Delete Icon)
      ================================================== */}
      {showBeforeSaleModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[70] p-2 sm:p-4">
          <div className="bg-white p-4 sm:p-5 rounded-2xl w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-3.5">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-sm sm:text-base font-bold text-slate-800">
                {editingReportId
                  ? "Edit Before Sale Report — "
                  : "Before Sale Report — "}
                {activeReportVisit.farmers?.farmer_name ||
                  activeReportVisit.batch_label}
              </h3>
              <button
                onClick={() => {
                  setEditingReportId(null);
                  setShowBeforeSaleModal(false);
                }}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 bg-emerald-50/70 border border-emerald-200 rounded-xl p-2.5 text-xs">
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Hatchery (DOC)
                </span>
                <strong className="text-emerald-900">
                  {activeReportVisit.hatchery || "-"}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  H.Date
                </span>
                <strong className="text-slate-800">
                  {activeReportVisit.hatch_date || "-"}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Total Birds
                </span>
                <strong className="text-slate-800">
                  {activeReportVisit.t_birds || 0} Pcs
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Age on Visit
                </span>
                <strong className="text-emerald-700">
                  {activeAgeDays} Days
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Visit Date
                </span>
                <input
                  type="date"
                  value={reportVisitDate}
                  min={activeReportVisit.hatch_date || undefined}
                  max={todayStr}
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="w-full border border-emerald-300 rounded px-1.5 py-0.5 bg-white text-xs font-semibold"
                />
              </div>
            </div>

            {checkingDate ? (
              <div className="py-10 text-center text-emerald-600 font-semibold text-xs">
                Checking date...
              </div>
            ) : existingReportOnDate ? (
              <div className="flex flex-col items-center justify-center p-8 bg-amber-50 border border-amber-200 rounded-xl text-center space-y-3">
                <div className="text-amber-900 font-bold text-sm">
                  Already a report on this date!
                </div>
                
                <button
                  type="button"
                  onClick={() => {
                    setShowBeforeSaleModal(false);
                    handleHistoryClick(activeReportVisit, activeReportVisit.farmers);
                  }}
                  className="mt-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg shadow-sm transition"
                >
                  Show the report from history
                </button>
              </div>
            ) : (
              <>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-xs">
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Mortality (Pcs) *
                    </label>
                    <input
                      type="number"
                      required
                      min={0}
                      max={reportMortalityCap > 0 ? reportMortalityCap : undefined}
                      value={beforeForm.mortality_pcs}
                      onChange={(e) =>
                        updateBeforeCalculations({ mortality_pcs: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Weight (g/b) *
                    </label>
                    <input
                      type="number"
                      required
                      value={beforeForm.weight_gm_b}
                      onChange={(e) =>
                        updateBeforeCalculations({ weight_gm_b: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-emerald-700 mb-1">
                      Std. Weight (Day {activeAgeDays})
                    </label>
                    <input
                      type="text"
                      disabled
                      value={
                        activeMeatStd?.weight ? `${activeMeatStd.weight} g` : ""
                      }
                      className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Total Feed (Kg) *
                    </label>
                    <input
                      type="number"
                      required
                      value={beforeForm.feed_t_kg}
                      onChange={(e) =>
                        updateBeforeCalculations({ feed_t_kg: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Feed/Bird (g) [Auto]
                    </label>
                    <input
                      type="text"
                      disabled
                      value={beforeForm.feed_gm_b}
                      className="w-full border rounded-lg px-2.5 py-1.5 bg-slate-100 font-bold"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-emerald-700 mb-1">
                      Std. Feed Intake
                    </label>
                    <input
                      type="text"
                      disabled
                      value={
                        activeMeatStd?.feed_intake
                          ? `${activeMeatStd.feed_intake} g`
                          : ""
                      }
                      className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      FCR (Without Mort.)
                    </label>
                    <input
                      type="text"
                      disabled
                      value={beforeForm.fcr_w_out_m}
                      className="w-full border rounded-lg px-2.5 py-1.5 bg-slate-100 font-bold text-blue-700"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-emerald-700 mb-1">
                      Std. FCR (W.out M)
                    </label>
                    <input
                      type="text"
                      disabled
                      value={activeMeatStd?.fcr ? `${activeMeatStd.fcr}` : ""}
                      className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      FCR (With Mort.)
                    </label>
                    <input
                      type="text"
                      disabled
                      value={beforeForm.fcr_with_m}
                      className="w-full border rounded-lg px-2.5 py-1.5 bg-slate-100 font-bold text-red-700"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-emerald-700 mb-1">
                      Feed Company *
                    </label>
                    <SearchableSelect
                      value={beforeForm.feed_company}
                      onChange={(v) => {
                        setBeforeForm({
                          ...beforeForm,
                          feed_company: v,
                        });
                        setTransferAgentId("");
                      }}
                      placeholder="Select Feed Company"
                      options={companies.map((c) => ({
                        value: c.company_name,
                        label: c.company_name,
                      }))}
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Feed Type *
                    </label>
                    <SearchableSelect
                    searchable={false}
                      value={beforeForm.feed_type}
                      onChange={(v) =>
                        setBeforeForm({ ...beforeForm, feed_type: v })
                      }
                      placeholder="Select Feed Type"
                      options={activeFeedTypeOptions.map((ft) => ({
                        value: ft,
                        label: ft,
                      }))}
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      MFG / Pro. Date
                    </label>
                    <input
                      type="date"
                      value={beforeForm.pro_date}
                      onChange={(e) =>
                        setBeforeForm({ ...beforeForm, pro_date: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Batch No.
                    </label>
                    <input
                      type="text"
                      value={beforeForm.batch_no}
                      onChange={(e) =>
                        setBeforeForm({ ...beforeForm, batch_no: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Indigestion (%)
                    </label>
                    <input
                      type="text"
                      placeholder="e.g. 5"
                      value={beforeForm.indigestion}
                      onChange={(e) =>
                        setBeforeForm({
                          ...beforeForm,
                          indigestion: e.target.value,
                        })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Orange Faeces (%)
                    </label>
                    <input
                      type="text"
                      placeholder="e.g. 5"
                      value={beforeForm.orange_feces}
                      onChange={(e) =>
                        setBeforeForm({
                          ...beforeForm,
                          orange_feces: e.target.value,
                        })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Watery Faeces (%)
                    </label>
                    <input
                      type="text"
                      placeholder="e.g. 5"
                      value={beforeForm.watery_feces}
                      onChange={(e) =>
                        setBeforeForm({
                          ...beforeForm,
                          watery_feces: e.target.value,
                        })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Uneven (%)
                    </label>
                    <input
                      type="text"
                      placeholder="e.g. 5"
                      value={beforeForm.uneven}
                      onChange={(e) =>
                        setBeforeForm({ ...beforeForm, uneven: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Litter Condition
                    </label>
                    <input
                      type="text"
                      value={beforeForm.litter_condition}
                      onChange={(e) =>
                        setBeforeForm({
                          ...beforeForm,
                          litter_condition: e.target.value,
                        })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>

                  {beforeChangedComp && (
                    <div className="col-span-2 sm:col-span-3 bg-amber-50 border border-amber-300 rounded-xl p-3 space-y-1.5">
                      <div className="text-xs font-bold text-amber-900">
                        Feed Company Changed to {beforeChangedComp.company_name}!
                        Select Agent under {beforeChangedComp.company_name}:
                      </div>
                      <select
                        value={transferAgentId}
                        onChange={(e) => setTransferAgentId(e.target.value)}
                        className="w-full border border-amber-400 rounded-lg px-3 py-1.5 bg-white text-xs font-semibold"
                      >
                        <option value="">-- Select Agent --</option>
                        {allAgents
                          .filter(
                            (a) =>
                              getAgentCompanyIds(a).includes(String(beforeChangedComp.id))
                          )
                          .map((ag) => (
                            <option key={ag.id} value={String(ag.id)}>
                              {ag.agent_name} ({ag.upazila})
                            </option>
                          ))}
                      </select>
                    </div>
                  )}

                  <div className="col-span-2 sm:col-span-3">
                    <label className="block font-bold text-slate-600 mb-1">
                      Remarks
                    </label>
                    <input
                      type="text"
                      placeholder="Write observation remarks..."
                      value={beforeForm.remarks}
                      onChange={(e) =>
                        setBeforeForm({ ...beforeForm, remarks: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                </div>

                
              </>
            )}

            <div className="flex flex-wrap justify-between items-center gap-2 pt-3 border-t">
              <div className="flex items-center gap-2">
                {editingReportId &&
                  canEditPoultryReport({ created_by: editingReportCreatedBy }) && (
                  <button
                    type="button"
                    title="Delete this report"
                    onClick={() =>
                      handleDeleteCurrentReport("before_sale_reports")
                    }
                    className="p-1.5 rounded-lg border border-red-200 bg-red-50 hover:bg-red-100 text-red-600 transition"
                  >
                    <Trash2 size={15} />
                  </button>
                )}
                <button
                  type="button"
                  onClick={() =>
                    generateBroilerWhatsAppText(
                      activeReportVisit,
                      beforeForm,
                      reportVisitDate,
                      true
                    )
                  }
                  className="px-3.5 py-1.5 rounded-lg bg-green-600 hover:bg-green-700 text-white text-xs font-bold shadow-sm"
                >
                  WhatsApp Msg
                </button>
              </div>

              <div className="flex gap-2">
                <button
                  onClick={() => {
                    setEditingReportId(null);
                    setShowBeforeSaleModal(false);
                  }}
                  className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold"
                >
                  Cancel
                </button>
                <button
                  onClick={handleBeforeSaleSubmit}
                  disabled={savingBefore}
                  className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold cursor-pointer disabled:opacity-60"
                >
                  {savingBefore
                    ? "Saving..."
                    : editingReportId
                    ? "Update Report"
                    : "Save Report"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 2: AFTER SALE REPORT (Supports both New & Edit from History + Delete Icon)
      ================================================== */}
      {showAfterSaleModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[70] p-2 sm:p-4">
          <div className="bg-white p-4 sm:p-5 rounded-2xl w-full max-w-lg max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-3.5">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-sm sm:text-base font-bold text-slate-800">
                {editingReportId
                  ? "Edit After Sale Report — "
                  : "After Sale Report — "}
                {activeReportVisit.farmers?.farmer_name ||
                  activeReportVisit.batch_label}
              </h3>
              <button
                onClick={() => {
                  setEditingReportId(null);
                  setShowAfterSaleModal(false);
                }}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 bg-red-50/70 border border-red-200 rounded-xl p-2.5 text-xs">
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Hatchery (DOC)
                </span>
                <strong className="text-red-900">
                  {activeReportVisit.hatchery || "-"}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  H.Date
                </span>
                <strong className="text-slate-800">
                  {activeReportVisit.hatch_date || "-"}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Total Birds
                </span>
                <strong className="text-slate-800">
                  {activeReportVisit.t_birds || 0} Pcs
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Age
                </span>
                <strong className="text-red-700">{activeAgeDays} Days</strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Sale Date
                </span>
                <input
                  type="date"
                  value={reportVisitDate}
                  min={activeReportVisit.hatch_date || undefined}
                  max={todayStr}
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="w-full border border-red-300 rounded px-1.5 py-0.5 bg-white text-xs font-semibold"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2.5 text-xs">
              <div className="col-span-2">
                <label className="block font-bold text-slate-600 mb-1">
                  Total Mortality (Pcs) *
                </label>
                <input
                  type="number"
                  required
                  min={0}
                  max={reportMortalityCap > 0 ? reportMortalityCap : undefined}
                  value={afterForm.mortality}
                  onChange={(e) =>
                    updateAfterCalculations({ mortality: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Avg Weight (g/bird) *
                </label>
                <input
                  type="number"
                  required
                  value={afterForm.weight}
                  onChange={(e) =>
                    updateAfterCalculations({ weight: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-emerald-700 mb-1">
                  Std. Weight (Day {activeAgeDays})
                </label>
                <input
                  type="text"
                  disabled
                  value={
                    activeMeatStd?.weight ? `${activeMeatStd.weight} g` : ""
                  }
                  className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                />
              </div>

              <div className="col-span-2">
                <label className="block font-bold text-slate-600 mb-1">
                  Total Feed Consumption (Kg) *
                </label>
                <input
                  type="number"
                  required
                  value={afterForm.feed_consumption_kg}
                  onChange={(e) =>
                    updateAfterCalculations({
                      feed_consumption_kg: e.target.value,
                    })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed/Bird (g) [Auto]
                </label>
                <input
                  type="text"
                  disabled
                  value={afterForm.feed_gm_b}
                  className="w-full border rounded-lg px-2.5 py-1.5 bg-slate-100 font-bold"
                />
              </div>
              <div>
                <label className="block font-bold text-emerald-700 mb-1">
                  Std. Feed Intake
                </label>
                <input
                  type="text"
                  disabled
                  value={
                    activeMeatStd?.feed_intake
                      ? `${activeMeatStd.feed_intake} g`
                      : ""
                  }
                  className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  FCR (Without Mort.)
                </label>
                <input
                  type="text"
                  disabled
                  value={afterForm.fcr_w_out_m}
                  className="w-full border rounded-lg px-2.5 py-1.5 bg-slate-100 font-bold text-blue-700"
                />
              </div>
              <div>
                <label className="block font-bold text-emerald-700 mb-1">
                  Std. FCR (W.out M)
                </label>
                <input
                  type="text"
                  disabled
                  value={activeMeatStd?.fcr ? `${activeMeatStd.fcr}` : ""}
                  className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                />
              </div>

              <div className="col-span-2">
                <label className="block font-bold text-slate-600 mb-1">
                  FCR (With Mortality)
                </label>
                <input
                  type="text"
                  disabled
                  value={afterForm.fcr_with_m}
                  className="w-full border rounded-lg px-2.5 py-1.5 bg-slate-100 font-bold text-red-700"
                />
              </div>

             <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Feed Company *
                    </label>
                    <SearchableSelect
                      value={afterForm.feed_mill}
                      onChange={(v) => {
                        setAfterForm({ ...afterForm, feed_mill: v });
                        setTransferAgentId("");
                      }}
                      placeholder="Select Feed Company"
                      options={companies.map((c) => ({
                        value: c.company_name,
                        label: c.company_name,
                      }))}
                    />
                  </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Last Feed *
                </label>
                <SearchableSelect
                searchable={false}
                  value={afterForm.feed}
                  onChange={(v) =>
                    setAfterForm({ ...afterForm, feed: v })
                  }
                  placeholder="Select Feed Type"
                  options={activeFeedTypeOptions.map((ft) => ({
                    value: ft,
                    label: ft,
                  }))}
                />
              </div>

              {afterChangedComp && (
                <div className="col-span-2 bg-amber-50 border border-amber-300 rounded-xl p-2.5 space-y-1.5">
                  <div className="text-xs font-bold text-amber-900">
                    Select Agent under {afterChangedComp.company_name}:
                  </div>
                  <select
                    value={transferAgentId}
                    onChange={(e) => setTransferAgentId(e.target.value)}
                    className="w-full border border-amber-400 rounded px-2 py-1 bg-white text-xs font-semibold"
                  >
                    <option value="">-- Select Agent --</option>
                    {allAgents
                      .filter(
                        (a) =>
                          getAgentCompanyIds(a).includes(String(afterChangedComp.id))
                      )
                      .map((ag) => (
                        <option key={ag.id} value={String(ag.id)}>
                          {ag.agent_name} ({ag.upazila})
                        </option>
                      ))}
                  </select>
                </div>
              )}

              <div className="col-span-2">
                <label className="block font-bold text-slate-600 mb-1">
                  Remarks
                </label>
                <input
                  type="text"
                  value={afterForm.remarks}
                  onChange={(e) =>
                    setAfterForm({ ...afterForm, remarks: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="flex justify-between items-center gap-2 pt-3 border-t">
              <div>
                {editingReportId &&
                  canEditPoultryReport({ created_by: editingReportCreatedBy }) && (
                  <button
                    type="button"
                    title="Delete this report"
                    onClick={() =>
                      handleDeleteCurrentReport("after_sale_reports")
                    }
                    className="p-1.5 rounded-lg border border-red-200 bg-red-50 hover:bg-red-100 text-red-600 transition"
                  >
                    <Trash2 size={15} />
                  </button>
                )}
              </div>
              <div className="flex gap-2">
                <button
                  onClick={() => {
                    setEditingReportId(null);
                    setShowAfterSaleModal(false);
                  }}
                  className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold"
                >
                  Cancel
                </button>
                <button
                  onClick={handleAfterSaleSoldSubmit}
                  disabled={savingAfter}
                  className="px-5 py-1.5 rounded-lg bg-red-600 hover:bg-red-700 text-white text-xs font-bold shadow-sm cursor-pointer disabled:opacity-60"
                >
                  {savingAfter
                    ? "Saving..."
                    : editingReportId
                    ? "Update Report"
                    : "Sold"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 3: LAYER PERFORMANCE REPORT (Supports both New & Edit from History + Delete Icon + Accurate White/Brown Std)
      ================================================== */}
      {showLayerModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[70] p-2 sm:p-4">
          <div className="bg-white p-4 sm:p-5 rounded-2xl w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-3.5">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-sm sm:text-base font-bold text-indigo-900">
                {editingReportId
                  ? "Edit Layer Performance Report — "
                  : "Layer Performance Report — "}
                {activeReportVisit.farmers?.farmer_name ||
                  activeReportVisit.batch_label}
              </h3>
              <button
                onClick={() => {
                  setEditingReportId(null);
                  setShowLayerModal(false);
                }}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-6 gap-2 bg-indigo-50/70 border border-indigo-200 rounded-xl p-2.5 text-xs">
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Hatchery (DOC)
                </span>
                <strong className="text-indigo-900">
                  {activeReportVisit.hatchery || layerForm.hatchery || "-"}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Breed
                </span>
                <strong className="text-indigo-900 truncate block">
                  {activeLayerBreed}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  H.Date
                </span>
                <strong className="text-slate-800">
                  {activeReportVisit.hatch_date || "-"}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Total Birds (Live)
                </span>
                <strong className="text-slate-800">
                  {liveLayerBirdsAfterMortality} Pcs
                  {((!editingReportId &&
                    Number(layerForm.mortality_pcs || 0) > 0) ||
                    (editingReportId &&
                      Number(layerForm.mortality_pcs || 0) !==
                        editingLayerOldMort)) && (
                      <span className="text-[10px] text-red-600 ml-1">
                        (
                        {editingReportId
                          ? `${activeReportVisit.t_birds} + ${editingLayerOldMort} - `
                          : `${activeReportVisit.t_birds} - `}
                        {Number(layerForm.mortality_pcs)})
                      </span>
                    )}
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Age (Weeks)
                </span>
                <strong className="text-indigo-700">
                  {activeAgeWeeks} Wks ({activeAgeDays}d)
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Visit Date
                </span>
                <input
                  type="date"
                  value={reportVisitDate}
                  min={activeReportVisit.hatch_date || undefined}
                  max={todayStr}
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="w-full border border-indigo-300 rounded px-1.5 py-0.5 bg-white text-xs font-semibold"
                />
              </div>
            </div>

            {checkingDate ? (
              <div className="py-10 text-center text-indigo-600 font-semibold text-xs">
                Checking date...
              </div>
            ) : existingReportOnDate ? (
              <div className="flex flex-col items-center justify-center p-8 bg-amber-50 border border-amber-200 rounded-xl text-center space-y-3">
                
                <div className="text-amber-900 font-bold text-sm">
                  Already a report on this date!
                </div>
                
                <button
                  type="button"
                  onClick={() => {
                    setShowLayerModal(false);
                    handleHistoryClick(activeReportVisit, activeReportVisit.farmers);
                  }}
                  className="mt-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg shadow-sm transition"
                >
                  Show the report from history
                </button>
              </div>
            ) : (
              <>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-xs">
                  <div>
                    <label className="block font-bold text-red-700 mb-1">
                      Mortality (Pcs) *
                    </label>
                    <input
                      type="number"
                      min={0}
                      max={layerMortalityCap > 0 ? layerMortalityCap : undefined}
                      required
                      value={layerForm.mortality_pcs}
                      onChange={(e) => {
                        const mVal = e.target.value;
                        setLayerForm((prev) => ({ ...prev, mortality_pcs: mVal }));
                        if (layerForm.feed_total_kg) {
                          handleLayerFeedChange(
                            "total_kg",
                            layerForm.feed_total_kg,
                            mVal
                          );
                        }
                      }}
                      className="w-full border border-red-300 rounded-lg px-2.5 py-1.5 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Body Weight (g) *
                    </label>
                    <input
                      type="number"
                      required
                      value={layerForm.b_weight_g}
                      onChange={(e) =>
                        setLayerForm({ ...layerForm, b_weight_g: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-indigo-700 mb-1">
                      Std. Weight (Week {activeAgeWeeks})
                    </label>
                    <input
                      type="text"
                      disabled
                      value={
                        activeLayerStd ? `${activeLayerStd.weight} g` : "N/A"
                      }
                      className="w-full border border-indigo-200 rounded-lg px-2.5 py-1.5 bg-indigo-50/50 text-indigo-900 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Feed Intake (Total Kg/Day) *
                    </label>
                    <input
                      type="number"
                      placeholder="Enter Total Kg"
                      value={layerForm.feed_total_kg}
                      onChange={(e) =>
                        handleLayerFeedChange("total_kg", e.target.value)
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Feed (gm/bird/day) *
                    </label>
                    <input
                      type="number"
                      placeholder="Or gm/bird/day"
                      value={layerForm.feed_intake_g}
                      onChange={(e) =>
                        handleLayerFeedChange("gm_bird", e.target.value)
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 font-bold"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-indigo-700 mb-1">
                      Std. Feed (gm/b/d)
                    </label>
                    <input
                      type="text"
                      disabled
                      value={
                        activeLayerStd
                          ? `${activeLayerStd.feed_bird_day} g`
                          : "N/A"
                      }
                      className="w-full border border-indigo-200 rounded-lg px-2.5 py-1.5 bg-indigo-50/50 text-indigo-900 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Egg Production (%) *
                    </label>
                    <input
                      type="number"
                      required
                      value={layerForm.egg_production}
                      onChange={(e) =>
                        setLayerForm({
                          ...layerForm,
                          egg_production: e.target.value,
                        })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div className="sm:col-span-2">
                    <label className="block font-bold text-indigo-700 mb-1">
                      Std. Egg Production (%)
                    </label>
                    <input
                      type="text"
                      disabled
                      value={
                        activeLayerStd
                          ? `${activeLayerStd.egg_production}%`
                          : "N/A"
                      }
                      className="w-full border border-indigo-200 rounded-lg px-2.5 py-1.5 bg-indigo-50/50 text-indigo-900 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Egg Weight (g) *
                    </label>
                    <input
                      type="number"
                      required
                      value={layerForm.egg_weight}
                      onChange={(e) =>
                        setLayerForm({ ...layerForm, egg_weight: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                  <div className="sm:col-span-2">
                    <label className="block font-bold text-indigo-700 mb-1">
                      Std. Egg Weight (g)
                    </label>
                    <input
                      type="text"
                      disabled
                      value={
                        activeLayerStd ? `${activeLayerStd.egg_weight} g` : "N/A"
                      }
                      className="w-full border border-indigo-200 rounded-lg px-2.5 py-1.5 bg-indigo-50/50 text-indigo-900 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-indigo-800 mb-1">
                      Feed Company *
                    </label>
                    <SearchableSelect
                      value={layerForm.feed_company}
                      onChange={(v) => {
                        setLayerForm({
                          ...layerForm,
                          feed_company: v,
                        });
                        setTransferAgentId("");
                      }}
                      placeholder="Select Feed Company"
                      options={companies.map((c) => ({
                        value: c.company_name,
                        label: c.company_name,
                      }))}
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Feed Type *
                    </label>
                    <SearchableSelect
                    searchable={false}
                      value={layerForm.feed_type}
                      onChange={(v) =>
                        setLayerForm({ ...layerForm, feed_type: v })
                      }
                      placeholder="Select Feed Type"
                      options={UPDATED_FEED_TYPES_BY_CATEGORY["Layer"].map((ft) => ({
                        value: ft,
                        label: ft,
                      }))}
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      MFG / Pro. Date
                    </label>
                    <input
                      type="date"
                      value={layerForm.pro_date}
                      onChange={(e) =>
                        setLayerForm({ ...layerForm, pro_date: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-600 mb-1">
                      Batch No.
                    </label>
                    <input
                      type="text"
                      value={layerForm.batch_no}
                      onChange={(e) =>
                        setLayerForm({ ...layerForm, batch_no: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>

                  <div className="sm:col-span-2">
                    <label className="block font-bold text-slate-600 mb-1">
                      Disease / Observation
                    </label>
                    <input
                      type="text"
                      value={layerForm.disease}
                      onChange={(e) =>
                        setLayerForm({ ...layerForm, disease: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>

                  {layerChangedComp && (
                    <div className="col-span-2 sm:col-span-3 bg-amber-50 border border-amber-300 rounded-xl p-2.5 space-y-1.5">
                      <div className="text-xs font-bold text-amber-900">
                        Select Agent under {layerChangedComp.company_name}:
                      </div>
                      <select
                        value={transferAgentId}
                        onChange={(e) => setTransferAgentId(e.target.value)}
                        className="w-full border border-amber-400 rounded px-2 py-1 bg-white text-xs font-semibold"
                      >
                        <option value="">-- Select Agent --</option>
                        {allAgents
                          .filter(
                            (a) =>
                              getAgentCompanyIds(a).includes(String(layerChangedComp.id))
                          )
                          .map((ag) => (
                            <option key={ag.id} value={String(ag.id)}>
                              {ag.agent_name} ({ag.upazila})
                            </option>
                          ))}
                      </select>
                    </div>
                  )}

                  <div className="col-span-2 sm:col-span-3">
                    <label className="block font-bold text-slate-600 mb-1">
                      Remarks
                    </label>
                    <input
                      type="text"
                      placeholder="Write observation remarks..."
                      value={layerForm.remarks}
                      onChange={(e) =>
                        setLayerForm({ ...layerForm, remarks: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                  </div>
                </div>

                {showCullReasonInput && (
                  <div className="bg-red-50 border border-red-300 rounded-xl p-3 space-y-2 text-xs">
                    <label className="block font-bold text-red-900">
                      Reason for Culling *
                    </label>
                    <div className="flex gap-2">
                      <input
                        type="text"
                        placeholder="Write reason why this batch is being culled..."
                        value={cullReasonText}
                        onChange={(e) => setCullReasonText(e.target.value)}
                        className="flex-1 border border-red-300 rounded-lg px-2.5 py-1.5 bg-white outline-none"
                      />
                      <button
                        type="button"
                        onClick={() => handleLayerReportSubmit(true)}
                        disabled={savingLayer}
                        className="px-4 py-1.5 rounded-lg bg-red-600 hover:bg-red-700 text-white font-bold shrink-0 shadow-sm cursor-pointer disabled:opacity-60"
                      >
                        {savingLayer ? "Saving..." : "Confirm Culled"}
                      </button>
                    </div>
                  </div>
                )}

                
              </>
            )}

            {showCullReasonInput && (
              <div className="bg-red-50 border border-red-300 rounded-xl p-3 space-y-2 text-xs">
                <label className="block font-bold text-red-900">
                  Reason for Culling *
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    placeholder="Write reason why this batch is being culled..."
                    value={cullReasonText}
                    onChange={(e) => setCullReasonText(e.target.value)}
                    className="flex-1 border border-red-300 rounded-lg px-2.5 py-1.5 bg-white outline-none"
                  />
                  <button
                    type="button"
                    onClick={() => handleLayerReportSubmit(true)}
                    disabled={savingLayer}
                    className="px-4 py-1.5 rounded-lg bg-red-600 hover:bg-red-700 text-white font-bold shrink-0 shadow-sm cursor-pointer disabled:opacity-60"
                  >
                    {savingLayer ? "Saving..." : "Confirm Culled"}
                  </button>
                </div>
              </div>
            )}

            <div className="flex flex-wrap justify-between items-center gap-2 pt-3 border-t">
              <div className="flex items-center gap-2">
                {editingReportId &&
                  canEditPoultryReport({ created_by: editingReportCreatedBy }) && (
                  <button
                    type="button"
                    title="Delete this report"
                    onClick={() => handleDeleteCurrentReport("layer_reports")}
                    className="p-1.5 rounded-lg border border-red-200 bg-red-50 hover:bg-red-100 text-red-600 transition"
                  >
                    <Trash2 size={15} />
                  </button>
                )}
                <button
                  type="button"
                  onClick={() =>
                    generateLayerWhatsAppText(
                      activeReportVisit,
                      layerForm,
                      true
                    )
                  }
                  className="px-3.5 py-1.5 rounded-lg bg-green-600 hover:bg-green-700 text-white text-xs font-bold shadow-sm"
                >
                  WhatsApp Msg
                </button>

                <button
                  type="button"
                  onClick={() => handleLayerReportSubmit(false)}
                  disabled={savingLayer}
                  className="px-4 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold cursor-pointer disabled:opacity-60"
                >
                  {savingLayer
                    ? "Saving..."
                    : editingReportId
                    ? "Update Layer Report"
                    : "Save Layer Report"}
                </button>

                
              </div>

              <div className="flex gap-2">
                  
                  {!editingReportId && (
                  <button
                    type="button"
                    disabled={showCullReasonInput}
                    onClick={() => setShowCullReasonInput(true)}
                    className={`px-3.5 py-1.5 rounded-lg text-xs font-bold shadow-sm transition ${
                      showCullReasonInput
                        ? "bg-slate-200 text-slate-400 cursor-not-allowed"
                        : "bg-red-600 hover:bg-red-700 text-white"
                    }`}
                  >
                    Culled
                  </button>
                )}

                <button
                  type="button"
                  onClick={() => {
                    setEditingReportId(null);
                    setShowLayerModal(false);
                  }}
                  className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold"
                >
                  Cancel
                </button>
                
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 4: MASTER VACCINE SCHEDULE
      ================================================== */}
      {showVaccineModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-3xl max-h-[90vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-4">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-base font-bold text-slate-800">
                Vaccine Schedule (Broiler, Layer, Sonali)
              </h3>
              <button
                onClick={() => setShowVaccineModal(false)}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            {isSuperAdmin && (
              <form
                onSubmit={handleAddVaccine}
                className="grid grid-cols-2 sm:grid-cols-6 gap-2 bg-slate-50 p-3 rounded-xl border border-slate-200 text-xs"
              >
                <select
                  value={vaccineForm.category}
                  onChange={(e) =>
                    setVaccineForm({ ...vaccineForm, category: e.target.value })
                  }
                  className="border rounded px-2 py-1.5 bg-white font-semibold"
                >
                  <option value="Broiler">Broiler</option>
                  <option value="Layer">Layer</option>
                  <option value="Sonali">Sonali</option>
                </select>

                <input
                  type="number"
                  required
                  placeholder="Age (Days)"
                  value={vaccineForm.age_days}
                  onChange={(e) =>
                    setVaccineForm({ ...vaccineForm, age_days: e.target.value })
                  }
                  className="border rounded px-2 py-1.5 bg-white"
                />

                <input
                  type="text"
                  required
                  placeholder="Vaccine Name"
                  value={vaccineForm.vaccine_name}
                  onChange={(e) =>
                    setVaccineForm({
                      ...vaccineForm,
                      vaccine_name: e.target.value,
                    })
                  }
                  className="sm:col-span-2 border rounded px-2 py-1.5 bg-white"
                />

                <input
                  type="text"
                  placeholder="Type (Live/Killed)"
                  value={vaccineForm.vaccine_type}
                  onChange={(e) =>
                    setVaccineForm({
                      ...vaccineForm,
                      vaccine_type: e.target.value,
                    })
                  }
                  className="border rounded px-2 py-1.5 bg-white"
                />

                <div className="flex gap-1">
                  <input
                    type="text"
                    placeholder="Route"
                    value={vaccineForm.route}
                    onChange={(e) =>
                      setVaccineForm({ ...vaccineForm, route: e.target.value })
                    }
                    className="w-full border rounded px-2 py-1.5 bg-white"
                  />
                  <button
                    type="submit"
                    disabled={savingAddVaccine}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 rounded font-bold shrink-0 cursor-pointer disabled:opacity-60"
                  >
                    {savingAddVaccine ? "Adding..." : "+ Add"}
                  </button>
                </div>
              </form>
            )}

            <div className="overflow-x-auto">
              <table className="w-full border-collapse border border-slate-300 text-xs">
                <thead className="bg-slate-100">
                  <tr>
                    <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                      Category
                    </th>
                    <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                      Age (Days)
                    </th>
                    <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                      Vaccine Name
                    </th>
                    <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                      Vaccine Type
                    </th>
                    <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                      Route
                    </th>
                    {isSuperAdmin && (
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                        Delete
                      </th>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {vaccines.map((v, i) => (
                    <tr key={v.id || i} className="hover:bg-slate-50">
                      <td className="border border-slate-300 px-2.5 py-1 font-semibold">
                        {v.category}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 text-center font-bold text-emerald-700">
                        Day {v.age_days}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 font-bold">
                        {v.vaccine_name}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {v.vaccine_type}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {v.route}
                      </td>
                      {isSuperAdmin && (
                        <td className="border border-slate-300 px-2.5 py-1 text-center">
                          <button
                            type="button"
                            onClick={() => handleDeleteVaccine(v.id)}
                            className="text-red-600 hover:underline font-semibold"
                          >
                            Delete
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 5: COMPACT BATCH-WISE HISTORY MODAL
          (Selected Batch upfront + Dropdown for other batches + Hatchery, H.Date & Breed in table)
      ================================================== */}
      {showHistoryModal && selectedHistoryVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-3">
          <div className="bg-white p-3.5 sm:p-4 rounded-xl w-full max-w-6xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-3">
            {/* Compact Header + Selected Batch upfront + Other Batches Dropdown on Side */}
            <div className="flex flex-wrap justify-between items-center gap-2 border-b pb-2">
              <div className="flex flex-wrap items-center gap-2">
                <h3 className="text-sm sm:text-base font-bold text-slate-800">
                  History:
                </h3>
                <span className="px-2.5 py-1 rounded-lg bg-emerald-600 text-white text-xs font-bold shadow-xs">
                  {histFarmerName} • {histBatchShort} (
                  {selectedHistoryVisit.category} •{" "}
                  {selectedHistoryVisit.status === "Activated"
                    ? "Running"
                    : selectedHistoryVisit.status}
                  )
                </span>

                {histShedShort && (
                  <span className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-300 text-slate-700 text-xs font-bold">
                    Shed: {histShedShort}
                  </span>
                )}

                {otherHistoryBatches.length > 0 && (
                  <select
                    value=""
                    onChange={(e) => {
                      const found = historyBatches.find(
                        (b) => String(b.id) === e.target.value
                      );
                      if (found) setSelectedHistoryVisit(found);
                    }}
                    className="border border-slate-300 rounded-lg px-2 py-1 bg-slate-50 hover:bg-slate-100 text-[11px] font-semibold text-slate-700 outline-none cursor-pointer"
                  >
                    <option value="">
                      Other Batches ({otherHistoryBatches.length}) ▾
                    </option>
                    {otherHistoryBatches.map((b: any) => {
                      const bFarmer =
                        b.farmers?.farmer_name || histFarmerName;
                      const bPrefix =
                        b.shedDisplayName &&
                        b.shedDisplayName !== bFarmer
                          ? b.shedDisplayName
                          : bFarmer;
                      return (
                        <option key={b.id} value={String(b.id)}>
                          {bPrefix} • {getHistoryBatchShort(b, bFarmer)} —{" "}
                          {b.category} (
                          {b.status === "Activated" ? "Running" : b.status} •{" "}
                          {b.hatch_date})
                        </option>
                      );
                    })}
                  </select>
                )}
              </div>

              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={exportVisitHistoryExcel}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded-lg text-[11px] font-semibold shadow-xs"
                >
                  Export Excel
                </button>
                <button
                  onClick={() => setShowHistoryModal(false)}
                  className="px-2 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold"
                >
                  ✖
                </button>
              </div>
            </div>

            {(selectedHistoryVisit.status === "Culled" ||
              selectedBatchCulledReason) && (
              <div className="bg-red-50 border border-red-200 rounded-lg px-2.5 py-1.5 text-[11px] text-red-900">
                <b>Culled Batch</b>
                {selectedHistoryVisit.closed_date
                  ? ` on ${formatDmy(selectedHistoryVisit.closed_date)}`
                  : ""}
                {selectedBatchCulledReason
                  ? ` — Reason: ${selectedBatchCulledReason}`
                  : ""}
              </div>
            )}

            {allReportsForFarmer.feedLogs.length > 0 && (
              <div className="bg-amber-50 border border-amber-200 rounded-lg px-2.5 py-1.5 text-[11px]">
                <strong className="text-amber-900">
                  Feed Conversion History:{" "}
                </strong>
                {allReportsForFarmer.feedLogs.map((l: any, idx: number) => (
                  <span key={idx} className="mr-3">
                    [{l.change_date}: {l.from_company} ➔ <b>{l.to_company}</b>]
                  </span>
                ))}
              </div>
            )}

            {(selectedHistoryVisit.category || "").toLowerCase() === "layer" ? (
              <div>
                <h4 className="text-[11px] font-bold uppercase text-indigo-800 mb-1.5">
                  Layer Performance History ({activeBatchLayerReports.length})
                </h4>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse border border-slate-300 text-[11px] leading-tight">
                    <thead className="bg-indigo-50 text-indigo-950 whitespace-nowrap">
                      <tr>
                        <th className="border border-slate-300 px-1.5 py-1">
                          Visit Date
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1">
                          Hatchery
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1">
                          H.Date
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1">
                          Breed
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1 text-center">
                          T.B
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1 text-center">
                          Age
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1 text-center">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1 text-center">
                          Feed / Std (g)
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1 text-center">
                          Egg% / Std%
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1 text-center">
                          Egg Wt / Std
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1">
                          Feed & Type
                        </th>
                        {layerExtraCols.map((col) => (
                          <th
                            key={col.key}
                            className="border border-slate-300 px-1.5 py-1"
                          >
                            {col.label}
                          </th>
                        ))}
                        <th className="border border-slate-300 px-1.5 py-1">
                          Remarks
                        </th>
                        <th className="border border-slate-300 px-1.5 py-1 text-center">
                          Action
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {activeBatchLayerReports.length === 0 ? (
                        <tr>
                          <td
                            colSpan={13 + layerExtraCols.length}
                            className="border border-slate-300 py-3 text-center text-slate-400"
                          >
                            No layer reports found for this batch.
                          </td>
                        </tr>
                      ) : (
                        activeBatchLayerReports.map((rep: any, i: number) => {
                          const wks = getAgeWeeks(
                            selectedHistoryVisit.hatch_date,
                            rep.visit_date
                          );
                          const breedVal =
                            selectedHistoryVisit.breed ||
                            rep.breed ||
                            "Shaver Star Cross 579 Brown";
                          const hatcheryVal =
                            selectedHistoryVisit.hatchery ||
                            rep.hatchery ||
                            "-";
                          const std = getLayerStdForVisit(breedVal, wks);
                          return (
                            <tr
                              key={i}
                              className="hover:bg-slate-50 whitespace-nowrap"
                            >
                              <td className="border border-slate-300 px-1.5 py-0.5">
                                {rep.visit_date || "-"}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5">
                                {hatcheryVal}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5">
                                {selectedHistoryVisit.hatch_date || "-"}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 font-medium text-indigo-900">
                                {breedVal}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                {selectedHistoryVisit.t_birds || 0}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 text-center font-bold">
                                {wks}w
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                <b>{rep.b_weight_g}</b> / {std?.weight || "-"}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                <b>{rep.feed_intake_g}</b> /{" "}
                                {std?.feed_bird_day || "-"}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                <b>{rep.egg_production}%</b> /{" "}
                                {std?.egg_production || "-"}%
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                <b>{rep.egg_weight}</b> /{" "}
                                {std?.egg_weight || "-"}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5">
                                {rep.feed_company || "-"}{" "}
                                {rep.feed_type ? `(${rep.feed_type})` : ""}
                              </td>
                              {layerExtraCols.map((col) => (
                                <td
                                  key={col.key}
                                  className="border border-slate-300 px-1.5 py-0.5"
                                >
                                  {hasValue(rep[col.key]) ? rep[col.key] : "-"}
                                </td>
                              ))}
                              <td className="border border-slate-300 px-1.5 py-0.5">
                                {getRemarkText(rep) || "-"}
                              </td>
                              <td className="border border-slate-300 px-1.5 py-0.5 text-center whitespace-nowrap">
                                <button
                                  type="button"
                                  onClick={() =>
                                    generateLayerWhatsAppText(
                                      selectedHistoryVisit,
                                      rep,
                                      false
                                    )
                                  }
                                  className="bg-green-600 text-white px-1.5 py-0.5 rounded text-[10px] font-bold"
                                >
                                  WtsApp
                                </button>
                                {canEditPoultryReport(rep) && (
                                <button
                                  type="button"
                                  onClick={() =>
                                    openEditReport("layer_reports", rep)
                                  }
                                  className="ml-1 bg-amber-500 text-white px-1.5 py-0.5 rounded text-[10px] font-bold"
                                >
                                  Edit
                                </button>
                                )}
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <>
                <div>
                  <h4 className="text-[11px] font-bold uppercase text-emerald-800 mb-1.5">
                    Before Sale Reports ({activeBatchBeforeReports.length})
                  </h4>
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse border border-slate-300 text-[11px] leading-tight">
                      <thead className="bg-slate-100 whitespace-nowrap">
                        <tr>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Visit Date
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Hatchery
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            H.Date
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Breed
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            T.B
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Age
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Mort.
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Wt / Std (g)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Feed (Kg)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Feed/B / Std (g)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            FCR(w/o) / Std
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            FCR(With)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Feed & Type
                          </th>
                          {beforeExtraCols.map((col) => (
                            <th
                              key={col.key}
                              className="border border-slate-300 px-1.5 py-1"
                            >
                              {col.label}
                            </th>
                          ))}
                          <th className="border border-slate-300 px-1.5 py-1">
                            Remarks
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Action
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {activeBatchBeforeReports.length === 0 ? (
                          <tr>
                            <td
                              colSpan={15 + beforeExtraCols.length}
                              className="border border-slate-300 py-3 text-center text-slate-400"
                            >
                              No before sale reports found for this batch.
                            </td>
                          </tr>
                        ) : (
                          activeBatchBeforeReports.map(
                            (rep: any, i: number) => {
                              const d = getAgeDays(
                                selectedHistoryVisit.hatch_date,
                                rep.visit_date
                              );
                              const std = getMeatStdForVisit(
                                selectedHistoryVisit.category,
                                d
                              );
                              return (
                                <tr
                                  key={i}
                                  className="hover:bg-slate-50 whitespace-nowrap"
                                >
                                  <td className="border border-slate-300 px-1.5 py-0.5">
                                    {rep.visit_date}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5">
                                    {selectedHistoryVisit.hatchery || "-"}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5">
                                    {selectedHistoryVisit.hatch_date || "-"}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5">
                                    {selectedHistoryVisit.breed || "-"}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                    {selectedHistoryVisit.t_birds}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center font-bold">
                                    {d}d
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                    {rep.mortality_pcs}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                    <b>{rep.weight_gm_b}</b> /{" "}
                                    {std?.weight || "-"}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                    {rep.feed_t_kg}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                    <b>{rep.feed_gm_b}</b> /{" "}
                                    {std?.feed_intake || "-"}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center font-bold">
                                    {rep.fcr_w_out_m} / {std?.fcr || "-"}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                    {rep.fcr_with_m}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5">
                                    {rep.feed_company || "-"}{" "}
                                    {rep.feed_type ? `(${rep.feed_type})` : ""}
                                  </td>
                                  {beforeExtraCols.map((col) => (
                                    <td
                                      key={col.key}
                                      className="border border-slate-300 px-1.5 py-0.5"
                                    >
                                      {hasValue(rep[col.key])
                                        ? rep[col.key]
                                        : "-"}
                                    </td>
                                  ))}
                                  <td className="border border-slate-300 px-1.5 py-0.5">
                                    {getRemarkText(rep) || "-"}
                                  </td>
                                  <td className="border border-slate-300 px-1.5 py-0.5 text-center whitespace-nowrap">
                                    <button
                                      type="button"
                                      onClick={() =>
                                        generateBroilerWhatsAppText(
                                          selectedHistoryVisit,
                                          rep,
                                          rep.visit_date,
                                          false
                                        )
                                      }
                                      className="bg-green-600 text-white px-1.5 py-0.5 rounded text-[10px] font-bold"
                                    >
                                      WtsApp
                                    </button>
                                    {canEditPoultryReport(rep) && (
                                    <button
                                      type="button"
                                      onClick={() =>
                                        openEditReport(
                                          "before_sale_reports",
                                          rep
                                        )
                                      }
                                      className="ml-1 bg-amber-500 text-white px-1.5 py-0.5 rounded text-[10px] font-bold"
                                    >
                                      Edit
                                    </button>
                                    )}
                                  </td>
                                </tr>
                              );
                            }
                          )
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div>
                  <h4 className="text-[11px] font-bold uppercase text-blue-800 mb-1.5">
                    After Sale Reports ({activeBatchAfterReports.length})
                  </h4>
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse border border-slate-300 text-[11px] leading-tight">
                      <thead className="bg-slate-100 whitespace-nowrap">
                        <tr>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Sale Date
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Hatchery
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            H.Date
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Breed
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            T.B
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Age
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Mort.
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Wt / Std (g)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Feed (Kg)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Feed/B / Std (g)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            FCR(w/o) / Std
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            FCR(With)
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Feed Company & Last Feed
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1">
                            Remarks
                          </th>
                          <th className="border border-slate-300 px-1.5 py-1 text-center">
                            Edit
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {activeBatchAfterReports.length === 0 ? (
                          <tr>
                            <td
                              colSpan={15}
                              className="border border-slate-300 py-3 text-center text-slate-400"
                            >
                              No after sale reports found for this batch.
                            </td>
                          </tr>
                        ) : (
                          activeBatchAfterReports.map((rep: any, i: number) => {
                            const d = getAgeDays(
                              selectedHistoryVisit.hatch_date,
                              rep.visit_date
                            );
                            const std = getMeatStdForVisit(
                              selectedHistoryVisit.category,
                              d
                            );
                            const tB = Number(
                              selectedHistoryVisit.t_birds || 0
                            );
                            const fGm =
                              rep.feed_gm_b ||
                              (tB > 0
                                ? (
                                    (Number(rep.feed_consumption_kg || 0) *
                                      1000) /
                                    tB
                                  ).toFixed(1)
                                : "-");
                            return (
                              <tr
                                key={i}
                                className="hover:bg-slate-50 whitespace-nowrap"
                              >
                                <td className="border border-slate-300 px-1.5 py-0.5">
                                  {rep.visit_date}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5">
                                  {selectedHistoryVisit.hatchery || "-"}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5">
                                  {selectedHistoryVisit.hatch_date || "-"}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5">
                                  {selectedHistoryVisit.breed || "-"}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                  {tB}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center font-bold">
                                  {d}d
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                  {rep.mortality}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                  <b>{rep.weight}</b> / {std?.weight || "-"}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                  {rep.feed_consumption_kg}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                  <b>{fGm}</b> / {std?.feed_intake || "-"}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center font-bold">
                                  {rep.fcr_w_out_m} / {std?.fcr || "-"}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                  {rep.fcr_with_m}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5">
                                  {rep.feed_mill || "-"}{" "}
                                  {rep.feed
                                    ? `(Last Feed: ${rep.feed})`
                                    : ""}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5">
                                  {getRemarkText(rep) || "-"}
                                </td>
                                <td className="border border-slate-300 px-1.5 py-0.5 text-center">
                                  {canEditPoultryReport(rep) && (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      openEditReport("after_sale_reports", rep)
                                    }
                                    className="bg-amber-500 text-white px-1.5 py-0.5 rounded text-[10px] font-bold"
                                  >
                                    Edit
                                  </button>
                                  )}
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 6: WHATSAPP MESSAGE PREVIEW & COPY POPUP
      ================================================== */}
      {whatsappPopupText && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[80] p-3">
          <div className="bg-white rounded-2xl p-5 w-full max-w-md shadow-2xl border border-slate-300 space-y-3">
            <div className="flex justify-between items-center border-b pb-2">
              <h3 className="text-sm font-bold text-slate-800">
                WhatsApp Observation Report
              </h3>
              <button
                onClick={() => setWhatsappPopupText("")}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <textarea
              readOnly
              rows={16}
              value={whatsappPopupText}
              className="w-full border border-slate-300 rounded-xl p-3 text-xs font-mono bg-slate-50 text-slate-800 outline-none"
            />

            <div className="flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setWhatsappPopupText("")}
                className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold"
              >
                Close
              </button>
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(whatsappPopupText);
                  setCopiedWa(true);
                }}
                className="px-4 py-1.5 rounded-lg bg-green-600 hover:bg-green-700 text-white text-xs font-bold shadow-sm"
              >
                {copiedWa ? "✓ Copied!" : "Copy Message"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}