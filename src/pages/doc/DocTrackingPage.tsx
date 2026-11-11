import React, { useEffect, useState, useRef, useMemo } from "react";
import XLSX from "xlsx-js-style";
import ExcelJS from "exceljs";
import { supabase } from "../../lib/supabase";
import LocationPicker from "../../components/maps/LocationPicker";
import {
  FaFileUpload,
  FaFileExcel,
  FaChartPie,
  FaPlus,
  FaTrash,
  FaEdit,
  FaTimes,
  FaSpinner,
  FaSyncAlt,
  FaCheck,
} from "react-icons/fa";
import {
  RotateCcw,
  MapPin,
  Navigation,
  CheckCircle2,
  X,
} from "lucide-react";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Capacitor } from "@capacitor/core";
import { Geolocation } from "@capacitor/geolocation";

const FARM_CATEGORIES = [
  "Broiler",
  "Layer",
  "Sonali",
  "Duck",
  "Quail",
  "Cattle",
  "Fish",
];

const BREEDS_BY_CATEGORY: Record<string, string[]> = {
  Broiler: ["Arbor Acress", "Hubbard EP", "Ross", "Cobb 500", "Indian River"],
  Layer: [
    "Shaver Star Cross 579 Brown",
    "Bovans White",
    "Shaver star cross 579",
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
    "Brown Cock",
    "Color Bird/Sasso",
    "White Cock",
    "Bornali",
    "Bornali Plus",
    "Local",
    "Classic",
    "Hybrid",
    "Dominant",
    "Hubbard Premium",
    "Sasso",
    "Tetra HB",
    "Cock",
  ],
  Duck: ["Khaki Campbell", "Pekin", "Jinding", "Deshi Duck"],
  Quail: ["Japanese Quail", "Bobwhite Quail"],
  Cattle: ["Holstein Friesian", "Sahiwal", "Sindhi", "Deshi", "Crossbred"],
  Fish: ["Pangus", "Tilapia", "Koi", "Pabda", "Gulsha", "Carp", "Mixed"],
};


function extractPcdCode(pcdRaw: string): string {
  const upper = String(pcdRaw || "").trim().toUpperCase();

  // Check exact match or distinct code token inside the string
  const knownCodes = ["AA", "EP", "HE", "RO", "SH", "BO", "HP", "SA", "BR", "BP", "CB", "CW"];
  for (const code of knownCodes) {
    if (upper === code) return code;
    const regex = new RegExp(`(^|[^A-Z])${code}([^A-Z]|$)`);
    if (regex.test(upper)) return code;
  }

  // Fallback prefix check
  for (const code of knownCodes) {
    if (upper.startsWith(code)) return code;
  }

  return upper;
}


// Updated Auto Category & Breed Mapping based on DOC Type (PCD)
function mapPcdToCategoryAndBreed(pcd: string): {
  category: string;
  breed: string;
} {
  const code = extractPcdCode(pcd);

  // 1. Broiler: AA, EP, HE, RO
  if (code === "AA") return { category: "Broiler", breed: "Arbor Acress" };
  if (code === "EP" || code === "HE")
    return { category: "Broiler", breed: "Hubbard EP" };
  if (code === "RO") return { category: "Broiler", breed: "Ross" };

  // 2. Layer: SH, BO
  if (code === "SH")
    return { category: "Layer", breed: "Shaver Star Cross 579 Brown" };
  if (code === "BO") return { category: "Layer", breed: "Bovans White" };

  // 3. Sonali: CB, HP, CW, BR, BP
  if (code === "CB") return { category: "Sonali", breed: "Brown Cock" };
  if (code === "HP" || code === "SA")
    return { category: "Sonali", breed: "Color Bird/Sasso" };
  if (code === "CW") return { category: "Sonali", breed: "White Cock" };
  if (code === "BP") return { category: "Sonali", breed: "Bornali Plus" };
  if (code === "BR") return { category: "Sonali", breed: "Bornali" };

  return { category: "Other", breed: String(pcd || "").trim() };
}

// Strict Category Filter Matcher (Broiler / Layer / Sonali)
const BROILER_CODES = ["AA", "EP", "HE", "RO"];
const LAYER_CODES = ["SH", "BO"];
const SONALI_CODES = ["HP", "SA", "BR", "BP", "CB", "CW"];

function matchesCategoryByPcd(pcd: string, categoryFilter: string): boolean {
  if (!categoryFilter) return true;
  const code = extractPcdCode(pcd);
  const cat = categoryFilter.trim().toLowerCase();

  if (cat === "broiler") return BROILER_CODES.includes(code);
  if (cat === "layer") return LAYER_CODES.includes(code);
  if (cat === "sonali") return SONALI_CODES.includes(code);

  return true;
}

function normStr(val: any): string {
  return String(val ?? "").trim();
}

function normDate(val: any): string {
  return String(val ?? "").trim().slice(0, 10);
}

function normCode(val: any): string {
  return String(val ?? "").trim().toUpperCase();
}

function normalizeFarmerStr(str?: string): string {
  return (str || "")
    .toLowerCase()
    .replace(
      /\b(md|mohammad|muhammad|mst|hazi|alhaj|mia|miah|sheikh|sk|dr)\b\.?/g,
      ""
    )
    .replace(/[^a-z0-9\u0980-\u09FF]/g, "")
    .trim();
}

const STOCK_KEY_BY_CATEGORY: Record<string, string> = {
  Broiler: "broiler",
  Layer: "layer",
  Sonali: "sonali",
  Duck: "duck",
  Quail: "quail",
};
const POULTRY_STOCK_KEYS = ["broiler", "layer", "sonali", "duck", "quail"];

interface DocAgent {
  id: string;
  doc_id: string;
  agent_name: string;
  proprietor: string;
  mobile: string;
  district: string;
  upazila: string;
  feed_agent_id?: number | null;
}

interface DocSupply {
  id: string;
  h_date: string;
  doc_id: string;
  agent_name: string;
  pcd: string;
  ag_b: string;
  qty: number;
  b_qty: number;
  rmk: string;
}

interface DocTracking {
  id: string;
  supply_id?: string | null;
  doc_id: string;
  h_date: string;
  pcd: string;
  farmer_name: string;
  address: string;
  mobile: string;
  qty: number;
  b_qty?: number | null;
  feed_company: string;
  remark: string;
  category?: string;
  breed?: string;
  ag_b?: string;
  farmer_id?: number | null;
  created_by?: string;
  tracked_by_name?: string;
}

interface OfficerRecord {
  id: string;
  full_name: string;
  email: string;
  role: string;
  upazilas: string[];
}

interface FeedAgentRecord {
  id: number;
  agent_code?: string;
  agent_name: string;
  upazila: string;
  company_id?: number;
}

interface FarmerRecord {
  id: number;
  agent_id: number;
  company_id: number;
  farmer_name: string;
  mobile: string;
  village: string;
  union_name: string;
  upazila: string;
  district?: string;
  latitude?: number | null;
  longitude?: number | null;
  farm_type?: string;
  farm_category?: string;
  broiler?: number;
  layer?: number;
  sonali?: number;
  duck?: number;
  quail?: number;
  companies?: { company_name?: string };
}

interface CompanyRecord {
  id: number;
  company_name: string;
}

let cachedPageState: {
  initialized: boolean;
  userId: string;
  userName: string;
  userRole: string;
  assignedUpazilas: string[];
  agents: DocAgent[];
  supplies: DocSupply[];
  trackings: DocTracking[];
  officers: OfficerRecord[];
  feedAgents: FeedAgentRecord[];
  farmersList: FarmerRecord[];
  companiesList: CompanyRecord[];
  companyNames: string[];
  fromDate: string;
  toDate: string;
  monthFilter: string;
  isMultiDate: boolean;
} | null = null;

export default function DocTrackingPage() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [loading, setLoading] = useState(!cachedPageState?.initialized);
  const [importing, setImporting] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");
  const [successMsg, setSuccessMsg] = useState("");

  const [currentUserId, setCurrentUserId] = useState(
    cachedPageState?.userId || ""
  );
  const [currentUserName, setCurrentUserName] = useState(
    cachedPageState?.userName || "User"
  );
  const [userRole, setUserRole] = useState(
    cachedPageState?.userRole || "officer"
  );
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>(
    cachedPageState?.assignedUpazilas || []
  );

  const [agents, setAgents] = useState<DocAgent[]>(
    cachedPageState?.agents || []
  );
  const [supplies, setSupplies] = useState<DocSupply[]>(
    cachedPageState?.supplies || []
  );
  const [trackings, setTrackings] = useState<DocTracking[]>(
    cachedPageState?.trackings || []
  );
  const [officers, setOfficers] = useState<OfficerRecord[]>(
    cachedPageState?.officers || []
  );
  const [feedAgents, setFeedAgents] = useState<FeedAgentRecord[]>(
    cachedPageState?.feedAgents || []
  );
  const [farmersList, setFarmersList] = useState<FarmerRecord[]>(
    cachedPageState?.farmersList || []
  );
  const [companiesList, setCompaniesList] = useState<CompanyRecord[]>(
    cachedPageState?.companiesList || []
  );
  const [companyNames, setCompanyNames] = useState<string[]>(
    cachedPageState?.companyNames || []
  );

  const todayStr = new Date().toISOString().split("T")[0];
  const [fromDate, setFromDate] = useState<string>(
    cachedPageState?.fromDate || ""
  );
  const [toDate, setToDate] = useState<string>(cachedPageState?.toDate || "");
  const [isMultiDate, setIsMultiDate] = useState<boolean>(
    cachedPageState?.isMultiDate || false
  );
  const [monthFilter, setMonthFilter] = useState<string>(
    cachedPageState?.monthFilter || ""
  );
  const [categoryFilter, setCategoryFilter] = useState<string>("");
  const [pcdFilter, setPcdFilter] = useState<string>("");
  const [upazilaFilter, setUpazilaFilter] = useState<string>("");
  const [agentDocIdFilter, setAgentDocIdFilter] = useState<string>("");
  const [officerFilter, setOfficerFilter] = useState<string>("");
  const [trackFilter, setTrackFilter] = useState<
    "all" | "tracked" | "untracked"
  >("all");

  // Modal Farmer Selection Mode:
  // "" (nothing selected initially) | "this_agent" | "other_agent" | "create_farmer" | "only_track"
  const [farmerEntryMode, setFarmerEntryMode] = useState<
    "" | "this_agent" | "other_agent" | "create_farmer" | "only_track"
  >("");

  // "Select from others agent" compact filters
  const [otherCompanyId, setOtherCompanyId] = useState("");
  const [otherUpazila, setOtherUpazila] = useState("");
  const [otherAgentId, setOtherAgentId] = useState("");

  // Floating Duplicate Suggestion state
  const [showSimilarSuggestions, setShowSimilarSuggestions] = useState(false);

  // Location Picker Modal inside "Create farmer & Track DOC"
  const [showCreateFarmerMap, setShowCreateFarmerMap] = useState(false);
  const [pickerCenterOverride, setPickerCenterOverride] = useState<
    [number, number] | null
  >(null);
  const [locatingInPicker, setLocatingInPicker] = useState(false);

  // Extra modal states
  const [saveToFarmers, setSaveToFarmers] = useState(false);
  const [showInVisits, setShowInVisits] = useState(true);
  const [batchMode, setBatchMode] = useState<"add" | "replace">("add");
  const [batchCtx, setBatchCtx] = useState<{
    farmerId: number;
    running: any[];
    prevCats: string[];
    inactive: boolean;
  } | null>(null);
  const [newFarmerForm, setNewFarmerForm] = useState({
    company_id: "",
    district: "",
    upazila: "",
    agent_id: "",
    union_name: "",
    village: "",
    latitude: "",
    longitude: "",
    feed_type: "",
    remarks: "",
  });
  const [geoDistricts, setGeoDistricts] = useState<any[]>([]);
  const [geoUpazilas, setGeoUpazilas] = useState<any[]>([]);
  const [geoUnions, setGeoUnions] = useState<any[]>([]);

  const [isTrackModalOpen, setIsTrackModalOpen] = useState(false);
  const [editingTracking, setEditingTracking] = useState<DocTracking | null>(
    null
  );
  const [selectedFarmerId, setSelectedFarmerId] = useState<string>("");

  const [activeTarget, setActiveTarget] = useState<{
    doc_id: string;
    agent_name: string;
    upazila: string;
    district: string;
    feed_agent_id?: number | null;
    pcd: string;
    h_date: string;
    supply_id?: string;
    remainingQty: number;
  } | null>(null);

  const [trackForm, setTrackForm] = useState({
    farmer_name: "",
    address: "",
    mobile: "",
    qty: "",
    box_qty: "",
    category: "Broiler",
    breed: "Arbor Acress",
    ag_b: "",
    h_date: "",
    pcd: "",
    batch_mark: "",
    feed_company: "",
    remark: "",
  });

  const [isExportModalOpen, setIsExportModalOpen] = useState(false);
  const [exportSelectedTypes, setExportSelectedTypes] = useState<string[]>([]);

  const [isReportOpen, setIsReportOpen] = useState(false);
  const [reportFromDate, setReportFromDate] = useState(todayStr);
  const [reportToDate, setReportToDate] = useState(todayStr);
  const [reportPcdFilter, setReportPcdFilter] = useState("");

  const isManagerOrAdmin =
    userRole === "super_admin" || userRole === "manager";

  useEffect(() => {
    if (cachedPageState?.initialized) {
      return;
    }
    initTrackingPage(true);
  }, []);

  function parseArray(field: any): string[] {
    if (!field) return [];
    if (Array.isArray(field)) return field;
    if (typeof field === "string") {
      try {
        const parsed = JSON.parse(field);
        if (Array.isArray(parsed)) return parsed;
      } catch {
        return field
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean);
      }
    }
    return [];
  }

  // Native & Browser Geolocation Helper for Create Farmer Location Picker
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

  async function handleLocateMeInNewFarmerPicker() {
    setLocatingInPicker(true);
    try {
      const { lat, lng } = await getDeviceCoordinates();
      setPickerCenterOverride([lat, lng]);
      setNewFarmerForm((prev) => ({
        ...prev,
        latitude: lat.toFixed(6),
        longitude: lng.toFixed(6),
      }));
    } catch (err: any) {
      alert(
        err?.message || "Unable to retrieve your location. Please allow GPS."
      );
    } finally {
      setLocatingInPicker(false);
    }
  }

  async function initTrackingPage(showSpinner = false) {
    if (showSpinner) setLoading(true);
    try {
      const {
        data: { user },
      } = await supabase.auth.getUser();

      let uid = user?.id || "";
      let uName = user?.email?.split("@")[0] || "Officer";
      let role = "officer";
      let uUpazilas: string[] = [];

      if (user) {
        const { data: profile } = await supabase
          .from("users")
          .select("id, full_name, email, role, upazilas")
          .eq("id", user.id)
          .single();

        if (profile) {
          uName = profile.full_name || profile.email || uName;
          role = profile.role || "officer";
          uUpazilas = parseArray(profile.upazilas);
        }
      }

      setCurrentUserId(uid);
      setCurrentUserName(uName);
      setUserRole(role);
      setAssignedUpazilas(uUpazilas);

      await loadAllData(role, uUpazilas, uid, uName);
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to initialize page.");
    } finally {
      if (showSpinner) setLoading(false);
    }
  }

  async function loadAllData(
    role = userRole,
    uUpazilas = assignedUpazilas,
    uid = currentUserId,
    uName = currentUserName
  ) {
    const isFullAccess = role === "super_admin" || role === "manager";

    const [
      { data: usersData },
      { data: faData },
      { data: frmData },
      { data: compData },
    ] = await Promise.all([
      supabase
        .from("users")
        .select("id, full_name, email, role, upazilas")
        .order("full_name"),
      supabase
        .from("agents")
        .select("id, agent_code, agent_name, upazila, company_id")
        .order("agent_name"),
      supabase
        .from("farmers")
        .select(
          "id, agent_id, company_id, farmer_name, mobile, village, union_name, upazila, district, latitude, longitude, farm_type, farm_category, broiler, layer, sonali, duck, quail, companies(company_name)"
        )
        .order("farmer_name"),
      supabase
        .from("companies")
        .select("id, company_name")
        .order("company_name"),
    ]);

    const parsedOfficers: OfficerRecord[] = (usersData || []).map((u: any) => ({
      id: u.id,
      full_name: u.full_name || u.email || "Unnamed",
      email: u.email || "",
      role: u.role || "officer",
      upazilas: parseArray(u.upazilas),
    }));

    const loadedComps: CompanyRecord[] = (compData as any[]) || [];
    const loadedCompNames: string[] = Array.from(
      new Set(loadedComps.map((c) => c.company_name?.trim()).filter(Boolean))
    );

    setOfficers(parsedOfficers);
    setFeedAgents((faData as any[]) || []);
    setFarmersList((frmData as any[]) || []);
    setCompaniesList(loadedComps);
    setCompanyNames(loadedCompNames);

    let agentQuery = supabase.from("doc_agents").select("*").order("doc_id");
    if (!isFullAccess) {
      if (uUpazilas.length > 0) {
        agentQuery = agentQuery.in("upazila", uUpazilas);
      } else {
        setAgents([]);
        setSupplies([]);
        setTrackings([]);
        return;
      }
    }

    const { data: agentData, error: agentErr } = await agentQuery;
    if (agentErr) throw agentErr;

    const loadedAgents: DocAgent[] = (agentData || []).map((a: any) => ({
      ...a,
      doc_id: normStr(a.doc_id),
    }));
    setAgents(loadedAgents);

    const allowedDocIds = loadedAgents.map((a) => a.doc_id);
    if (allowedDocIds.length === 0) {
      setSupplies([]);
      setTrackings([]);
      return;
    }

    const [{ data: supData, error: supErr }, { data: trkData, error: trkErr }] =
      await Promise.all([
        supabase
          .from("doc_supplies")
          .select("*")
          .in("doc_id", allowedDocIds)
          .order("h_date", { ascending: false }),
        supabase
          .from("doc_trackings")
          .select("*")
          .in("doc_id", allowedDocIds)
          .order("created_at", { ascending: true }),
      ]);

    if (supErr) throw supErr;
    if (trkErr) throw trkErr;

    const loadedSupplies: DocSupply[] = (supData || []).map((s: any) => ({
      ...s,
      doc_id: normStr(s.doc_id),
      h_date: normDate(s.h_date),
      pcd: normCode(s.pcd),
    }));

    const loadedTrackings: DocTracking[] = (trkData || []).map((t: any) => {
      let trackerName = t.tracked_by_name;
      if (!trackerName && t.created_by) {
        const foundOff = parsedOfficers.find((o) => o.id === t.created_by);
        if (foundOff) trackerName = foundOff.full_name;
      }
      return {
        ...t,
        doc_id: normStr(t.doc_id),
        h_date: normDate(t.h_date),
        pcd: normCode(t.pcd),
        tracked_by_name: trackerName || "-",
      };
    });

    setSupplies(loadedSupplies);
    setTrackings(loadedTrackings);

    let defFrom = fromDate;
    let defTo = toDate;
    if (loadedSupplies.length > 0 && !defFrom && !defTo && !monthFilter) {
      defFrom = loadedSupplies[0].h_date;
      defTo = loadedSupplies[0].h_date;
      setFromDate(defFrom);
      setToDate(defTo);
      setReportFromDate(defFrom);
      setReportToDate(defTo);
    }

    cachedPageState = {
      initialized: true,
      userId: uid,
      userName: uName,
      userRole: role,
      assignedUpazilas: uUpazilas,
      agents: loadedAgents,
      supplies: loadedSupplies,
      trackings: loadedTrackings,
      officers: parsedOfficers,
      feedAgents: (faData as any[]) || [],
      farmersList: (frmData as any[]) || [],
      companiesList: loadedComps,
      companyNames: loadedCompNames,
      fromDate: defFrom,
      toDate: defTo,
      monthFilter,
      isMultiDate,
    };
  }

  function extractCleanDocId(rawVal: any): string {
    if (rawVal === null || rawVal === undefined) return "";
    return String(rawVal).replace(/<[^>]*>/g, "").trim();
  }

  function parseExcelDate(rawDate: any): string {
    if (!rawDate) return todayStr;
    if (typeof rawDate === "number") {
      const utcDays = Math.floor(rawDate - 25569);
      const dateObj = new Date(utcDays * 86400 * 1000);
      return dateObj.toISOString().split("T")[0];
    }
    const str = String(rawDate).replace(/<[^>]*>/g, "").trim();
    const dmyMatch = str.match(
      /^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/
    );
    if (dmyMatch) {
      const day = dmyMatch[1].padStart(2, "0");
      const month = dmyMatch[2].padStart(2, "0");
      let year = dmyMatch[3];
      if (year.length === 2) year = "20" + year;
      return `${year}-${month}-${day}`;
    }
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString().split("T")[0];
    }
    return todayStr;
  }

  async function handleExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    setImporting(true);
    setErrorMsg("");
    setSuccessMsg("");

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (rows.length === 0) {
        throw new Error("The selected Excel file is empty.");
      }

      const groupedMap = new Map<
        string,
        {
          h_date: string;
          doc_id: string;
          agent_name: string;
          pcd: string;
          ag_b: string;
          qty: number;
          b_qty: number;
          rmk: string;
        }
      >();

      for (const row of rows) {
        const getVal = (possibleKeys: string[]) => {
          for (const k of Object.keys(row)) {
            const cleanKey = k.trim().toLowerCase();
            if (possibleKeys.includes(cleanKey)) return row[k];
          }
          return "";
        };

        const rawDocId = getVal(["doc.id", "doc id", "doc_id", "docid"]);
        const cleanDocId = extractCleanDocId(rawDocId);
        if (!cleanDocId) continue;

        const rawHDate = getVal(["h.date", "h date", "h_date", "date"]);
        const cleanDate = parseExcelDate(rawHDate);

        const rawAgent = String(getVal(["agent", "agent name"]) || "")
          .replace(/<[^>]*>/g, "")
          .trim();

        const rawPcd = normCode(
          String(getVal(["pcd", "doc type"]) || "").replace(/<[^>]*>/g, "")
        );
        if (!rawPcd) continue;

        const rawAgB = String(getVal(["ag/b+", "ag/b", "ag_b"]) || "")
          .replace(/<[^>]*>/g, "")
          .trim();

        const qtyVal =
          Number(
            String(getVal(["qty", "quantity"]) || "0").replace(/[^0-9.-]/g, "")
          ) || 0;

        const bQtyVal =
          Number(
            String(getVal(["b.qty", "b qty", "b_qty"]) || "0").replace(
              /[^0-9.-]/g,
              ""
            )
          ) || 0;

        const rawRmk = String(getVal(["rmk", "remark", "remarks"]) || "")
          .replace(/<[^>]*>/g, "")
          .trim();

        const uniqueKey = `${cleanDate}_${cleanDocId}_${rawPcd}_${rawAgB}`;

        if (groupedMap.has(uniqueKey)) {
          const existing = groupedMap.get(uniqueKey)!;
          existing.qty += qtyVal;
          existing.b_qty += bQtyVal;
        } else {
          groupedMap.set(uniqueKey, {
            h_date: cleanDate,
            doc_id: cleanDocId,
            agent_name: rawAgent,
            pcd: rawPcd,
            ag_b: rawAgB,
            qty: qtyVal,
            b_qty: bQtyVal,
            rmk: rawRmk,
          });
        }
      }

      const payload = Array.from(groupedMap.values());
      if (payload.length === 0) {
        throw new Error(
          "No valid DOC rows found. Check columns: H.Date, DOC.ID, PCD, Qty."
        );
      }

      const importedDates = Array.from(new Set(payload.map((p) => p.h_date)));
      const importedDocIds = Array.from(new Set(payload.map((p) => p.doc_id)));

      await supabase
        .from("doc_supplies")
        .delete()
        .in("h_date", importedDates)
        .in("doc_id", importedDocIds);

      const { error: insErr } = await supabase
        .from("doc_supplies")
        .insert(payload);

      if (insErr) throw insErr;

      const importedDate = payload[0]?.h_date;
      if (importedDate) {
        setFromDate(importedDate);
        setToDate(importedDate);
        setIsMultiDate(false);
        setMonthFilter("");
        if (cachedPageState) {
          cachedPageState.fromDate = importedDate;
          cachedPageState.toDate = importedDate;
          cachedPageState.isMultiDate = false;
          cachedPageState.monthFilter = "";
        }
      }

      setSuccessMsg(
        `Imported ${payload.length} DOC supply rows for H.Date: ${importedDate}!`
      );
      await loadAllData();
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to import Excel file.");
    } finally {
      setImporting(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }

  function getCollaboratedFeedAgentId(
    target: typeof activeTarget
  ): number | null {
    if (!target) return feedAgents[0]?.id || null;
    if (target.feed_agent_id) return Number(target.feed_agent_id);

    const byCode = feedAgents.find(
      (fa) => fa.agent_code && fa.agent_code.trim() === target.doc_id.trim()
    );
    if (byCode) return byCode.id;

    const byName = feedAgents.find(
      (fa) =>
        fa.agent_name.trim().toLowerCase() ===
        target.agent_name.trim().toLowerCase()
    );
    if (byName) return byName.id;

    const inUpazila = feedAgents.find((fa) => fa.upazila === target.upazila);
    if (inUpazila) return inUpazila.id;

    return feedAgents[0]?.id || null;
  }

  function getCollaboratedFarmersForTarget(target: typeof activeTarget) {
    if (!target) return [];

    let matchedFeedAgentIds: number[] = [];
    if (target.feed_agent_id) {
      matchedFeedAgentIds.push(Number(target.feed_agent_id));
    } else {
      const codeOrNameMatches = feedAgents.filter(
        (fa) =>
          (fa.agent_code && fa.agent_code.trim() === target.doc_id.trim()) ||
          fa.agent_name.trim().toLowerCase() ===
            target.agent_name.trim().toLowerCase()
      );
      if (codeOrNameMatches.length > 0) {
        matchedFeedAgentIds = codeOrNameMatches.map((fa) => fa.id);
      }
    }

    if (matchedFeedAgentIds.length > 0) {
      const underFeedAgent = farmersList.filter((f) =>
        matchedFeedAgentIds.includes(Number(f.agent_id))
      );
      if (underFeedAgent.length > 0) return underFeedAgent;
    }

    return farmersList.filter((f) => f.upazila === target.upazila);
  }

  function canModifyTracking(tr: DocTracking): boolean {
    if (isManagerOrAdmin) return true;
    if (tr.created_by && currentUserId && tr.created_by === currentUserId) {
      return true;
    }
    if (
      !tr.created_by &&
      tr.tracked_by_name &&
      tr.tracked_by_name === currentUserName
    ) {
      return true;
    }
    return false;
  }

  function getStatsForAgentDatePcd(
    docId: string,
    hDate: string,
    pcd: string,
    excludeTrackingId?: string
  ) {
    const cleanId = normStr(docId);
    const cleanDate = normDate(hDate);
    const cleanPcd = normCode(pcd);

    const supRows = supplies.filter(
      (s) =>
        normStr(s.doc_id) === cleanId &&
        normDate(s.h_date) === cleanDate &&
        normCode(s.pcd) === cleanPcd
    );
    const total = supRows.reduce((sum, r) => sum + (Number(r.qty) || 0), 0);
    const bTotal = supRows.reduce((sum, r) => sum + (Number(r.b_qty) || 0), 0);
    const docPerBox = total > 0 && bTotal > 0 ? total / bTotal : 0;

    const trkRows = trackings.filter(
      (t) =>
        normStr(t.doc_id) === cleanId &&
        normDate(t.h_date) === cleanDate &&
        normCode(t.pcd) === cleanPcd &&
        (!excludeTrackingId || String(t.id) !== String(excludeTrackingId))
    );
    const tracked = trkRows.reduce((sum, r) => sum + (Number(r.qty) || 0), 0);
    const remaining = Math.max(0, total - tracked);

    const agBOptions = Array.from(
      new Set(supRows.map((s) => s.ag_b).filter(Boolean))
    );

    return {
      total,
      bTotal,
      docPerBox,
      tracked,
      remaining,
      supplyId: supRows[0]?.id,
      agBOptions: agBOptions.length > 0 ? agBOptions : ["AG", "B+"],
    };
  }

  function computeAutoBoxQty(
    docQtyVal: string | number,
    docPerBox: number
  ): string {
    const q = Number(docQtyVal);
    if (!q || q <= 0 || !docPerBox || docPerBox <= 0) return "";
    return String(Math.round(q / docPerBox));
  }

  // Single "Add Tracking" button handler per agent
  function openAgentTrackingModal(agent: DocAgent, agentBlocks: any[]) {
    const firstUntracked =
      agentBlocks.find((b) => b.untracked > 0) || agentBlocks[0];
    if (!firstUntracked || firstUntracked.untracked <= 0) {
      alert("All DOC are tracked for this agent.");
      return;
    }

    const hDate = normDate(firstUntracked.h_date);
    const pcd = normCode(firstUntracked.pcd);
    const stats = getStatsForAgentDatePcd(agent.doc_id, hDate, pcd);

    const defaultFeedComp =
      companyNames.find((c) => c.toLowerCase().includes("nourish")) ||
      companyNames[0] ||
      "";

    const { category: autoCat, breed: autoBreed } =
      mapPcdToCategoryAndBreed(pcd);

    const autoBoxes = computeAutoBoxQty(stats.remaining, stats.docPerBox);

    setEditingTracking(null);
    setFarmerEntryMode(""); // Start clean with no form shown until an option is selected
    setOtherCompanyId("");
    setOtherUpazila(agent.upazila || "");
    setOtherAgentId("");
    setShowSimilarSuggestions(false);
    setShowCreateFarmerMap(false);
    setPickerCenterOverride(null);

    setActiveTarget({
      doc_id: agent.doc_id,
      agent_name: agent.agent_name,
      upazila: agent.upazila,
      district: agent.district,
      feed_agent_id: agent.feed_agent_id,
      pcd,
      h_date: hDate,
      supply_id: stats.supplyId,
      remainingQty: stats.remaining,
    });
    setSelectedFarmerId("");
    setSaveToFarmers(false);
    setShowInVisits(true);
    setBatchMode("add");
    setBatchCtx(null);
    setTrackForm({
      farmer_name: "",
      address: "",
      mobile: "",
      qty: String(stats.remaining),
      box_qty: autoBoxes,
      category: autoCat,
      breed: autoBreed,
      ag_b: stats.agBOptions.length === 1 ? stats.agBOptions[0] : "",
      h_date: hDate,
      pcd,
      batch_mark: "",
      feed_company: defaultFeedComp,
      remark: "",
    });
    setIsTrackModalOpen(true);
  }

  function openEditTrackingModal(tr: DocTracking, agent: DocAgent) {
    if (!canModifyTracking(tr)) {
      alert("You can only edit tracking entries added by you.");
      return;
    }

    const cleanDate = normDate(tr.h_date);
    const cleanPcd = normCode(tr.pcd);
    const stats = getStatsForAgentDatePcd(
      agent.doc_id,
      cleanDate,
      cleanPcd,
      tr.id
    );
    const { category: autoCat, breed: autoBreed } =
      mapPcdToCategoryAndBreed(cleanPcd);

    const initialBoxQty =
      tr.b_qty !== undefined && tr.b_qty !== null
        ? String(tr.b_qty)
        : computeAutoBoxQty(tr.qty, stats.docPerBox);

    setEditingTracking(tr);
    setFarmerEntryMode("this_agent");
    setShowSimilarSuggestions(false);
    setActiveTarget({
      doc_id: agent.doc_id,
      agent_name: agent.agent_name,
      upazila: agent.upazila,
      district: agent.district,
      feed_agent_id: agent.feed_agent_id,
      pcd: cleanPcd,
      h_date: cleanDate,
      supply_id: tr.supply_id || stats.supplyId,
      remainingQty: stats.remaining,
    });
    setSelectedFarmerId(tr.farmer_id ? String(tr.farmer_id) : "");
    setTrackForm({
      farmer_name: tr.farmer_name || "",
      address: tr.address || "",
      mobile: tr.mobile || "",
      qty: String(tr.qty || ""),
      box_qty: initialBoxQty,
      category: tr.category || autoCat,
      breed: tr.breed || autoBreed,
      ag_b: tr.ag_b || "",
      h_date: cleanDate,
      pcd: cleanPcd,
      batch_mark: "",
      feed_company: tr.feed_company || companyNames[0] || "",
      remark: tr.remark || "",
    });
    setSaveToFarmers(false);
    setShowInVisits(false);
    setBatchCtx(null);
    setIsTrackModalOpen(true);
  }

  function populateFormFromFarmerRecord(f: FarmerRecord) {
    const compName = f.companies?.company_name || "";
    const matchedFeedComp =
      companyNames.find((c) => c.toLowerCase() === compName.toLowerCase()) ||
      trackForm.feed_company ||
      companyNames[0] ||
      "";

    const addr = [f.village, f.union_name, f.upazila]
      .filter(Boolean)
      .join(", ");

    const mapped = mapPcdToCategoryAndBreed(
      trackForm.pcd || activeTarget?.pcd || ""
    );

    setTrackForm((prev) => ({
      ...prev,
      farmer_name: f.farmer_name || "",
      address: addr,
      mobile: f.mobile || "",
      feed_company: matchedFeedComp,
      category: mapped.category,
      breed: mapped.breed,
    }));
  }

  function handleSelectFarmerFromDropdown(farmerId: string) {
    setSelectedFarmerId(farmerId);
    if (!farmerId) return;

    const f = farmersList.find((item) => String(item.id) === String(farmerId));
    if (!f) return;
    populateFormFromFarmerRecord(f);
  }

  // Pick a farmer from Floating Duplicate Suggestions
  function handlePickSuggestedFarmer(f: FarmerRecord) {
    setSelectedFarmerId(String(f.id));
    setSaveToFarmers(false);
    setShowInVisits(true);
    setShowSimilarSuggestions(false);

    const currentFeedAgentId = getCollaboratedFeedAgentId(activeTarget);
    if (
      currentFeedAgentId &&
      Number(f.agent_id) === Number(currentFeedAgentId)
    ) {
      setFarmerEntryMode("this_agent");
    } else {
      setFarmerEntryMode("other_agent");
      setOtherCompanyId(f.company_id ? String(f.company_id) : "");
      setOtherUpazila(f.upazila || "");
      setOtherAgentId(f.agent_id ? String(f.agent_id) : "");
    }

    populateFormFromFarmerRecord(f);
  }

  async function handleSaveTracking(e: React.FormEvent) {
    e.preventDefault();
    if (!activeTarget) return;

    const trackHDate = normDate(trackForm.h_date || activeTarget.h_date);
    const trackPcd = normCode(trackForm.pcd || activeTarget.pcd);
    const liveStats = getStatsForAgentDatePcd(
      activeTarget.doc_id,
      trackHDate,
      trackPcd,
      editingTracking?.id
    );

    const numQty = Number(trackForm.qty);
    const numBoxes =
      trackForm.box_qty !== ""
        ? Math.max(0, Number(trackForm.box_qty) || 0)
        : liveStats.docPerBox > 0
        ? Math.round(numQty / liveStats.docPerBox)
        : 0;

    // 1 bonus DOC per box is added ONLY for Farmers page stock & Vet Visits Running page
    const effectiveBirdsForFarmAndVisit = numQty + numBoxes;

    if (!trackForm.farmer_name.trim() || numQty <= 0) {
      setErrorMsg("Farmer Name and valid DOC Qty are required.");
      return;
    }

    if (numQty > liveStats.remaining) {
      alert(
        `DOC Qty (${numQty}) cannot exceed remaining untracked DOC (${liveStats.remaining}).`
      );
      return;
    }

    setSubmitting(true);
    setErrorMsg("");

    try {
      const finalCategory = trackForm.category || "Broiler";
      const finalBreed = trackForm.breed || "Arbor Acress";

      if (editingTracking) {
        if (!canModifyTracking(editingTracking)) {
          throw new Error("Permission denied to edit this tracking.");
        }

        const oldQty = Number(editingTracking.qty || 0);
        const qtyDiff = numQty - oldQty;

        const updateData: Record<string, any> = {
          farmer_name: trackForm.farmer_name.trim(),
          address: trackForm.address.trim(),
          mobile: trackForm.mobile.trim(),
          qty: numQty,
          b_qty: numBoxes,
          feed_company: trackForm.feed_company,
          remark: trackForm.remark.trim(),
          category: finalCategory,
          breed: finalBreed,
          ag_b: trackForm.ag_b || null,
        };

        const { error } = await supabase
          .from("doc_trackings")
          .update(updateData)
          .eq("id", editingTracking.id);

        if (error) {
          delete updateData.b_qty;
          delete updateData.category;
          delete updateData.breed;
          delete updateData.ag_b;
          const { error: retryErr } = await supabase
            .from("doc_trackings")
            .update(updateData)
            .eq("id", editingTracking.id);
          if (retryErr) throw retryErr;
        }

        // --- Farmer Page & Vet Visits Stock Sync on Edit ---
        const targetFarmerId = editingTracking.farmer_id || selectedFarmerId;
        if (targetFarmerId && qtyDiff !== 0) {
          const { data: frmRec } = await supabase
            .from("farmers")
            .select("*")
            .eq("id", targetFarmerId)
            .single();

          if (frmRec) {
            const stockKey = STOCK_KEY_BY_CATEGORY[finalCategory];
            if (stockKey) {
              const currentStock = Number(frmRec[stockKey] || 0);
              const boxDiff = numBoxes - Number(editingTracking.b_qty || 0);
              const effectiveDiff = qtyDiff + boxDiff;

              const newStock = Math.max(0, currentStock + effectiveDiff);
              await supabase
                .from("farmers")
                .update({ [stockKey]: newStock })
                .eq("id", targetFarmerId);
            }
          }

          const { data: runningVisit } = await supabase
            .from("vet_visits")
            .select("id, t_birds")
            .eq("farmer_id", String(targetFarmerId))
            .in("status", ["Activated", "Running"])
            .order("created_at", { ascending: false })
            .limit(1)
            .single();

          if (runningVisit) {
            const currentVisitBirds = Number(runningVisit.t_birds || 0);
            const boxDiff = numBoxes - Number(editingTracking.b_qty || 0);
            const effectiveVisitDiff = qtyDiff + boxDiff;
            const newVisitBirds = Math.max(
              0,
              currentVisitBirds + effectiveVisitDiff
            );

            await supabase
              .from("vet_visits")
              .update({
                t_birds: newVisitBirds,
                box_no: numBoxes > 0 ? String(numBoxes) : "",
              })
              .eq("id", runningVisit.id);
          }
        }

        setTrackings((prev) =>
          prev.map((item) =>
            String(item.id) === String(editingTracking.id)
              ? {
                  ...item,
                  ...updateData,
                  b_qty: numBoxes,
                  ag_b: trackForm.ag_b || "",
                }
              : item
          )
        );
      } else {
        const isOnlyTrackMode = farmerEntryMode === "only_track";
        const isCreateFarmerMode = farmerEntryMode === "create_farmer";
        const shouldSaveFarmer = !isOnlyTrackMode && (isCreateFarmerMode || saveToFarmers);
        const shouldShowInVisits = !isOnlyTrackMode && showInVisits;

        const feedAgentFromTarget = getCollaboratedFeedAgentId(activeTarget);
        const chosenCompObj =
          companiesList.find(
            (c) =>
              c.company_name?.trim().toLowerCase() ===
              trackForm.feed_company.trim().toLowerCase()
          ) ||
          companiesList.find((c) =>
            c.company_name?.toLowerCase().includes("nourish")
          ) ||
          companiesList[0];
        const targetCompId = chosenCompObj?.id || 1;

        // 1. Identify or create Farmer record (Skipped completely if "Only Track" is selected)
        let targetFarmerId: number | null =
          !isOnlyTrackMode && selectedFarmerId
            ? Number(selectedFarmerId)
            : null;

        const stockKey = STOCK_KEY_BY_CATEGORY[finalCategory];

        // 2. If user selected "Create farmer & Track DOC"
        if (!isOnlyTrackMode && !targetFarmerId && shouldSaveFarmer) {
          const nfAgent = Number(
            newFarmerForm.agent_id || feedAgentFromTarget || 0
          );
          const nfComp = Number(newFarmerForm.company_id || targetCompId);

          if (!nfAgent) {
            throw new Error(
              "Please select an Agent to create this farmer in the Farmers database."
            );
          }

          const newFarmerPayload: Record<string, any> = {
            company_id: nfComp,
            agent_id: nfAgent,
            farmer_name: trackForm.farmer_name.trim(),
            mobile: trackForm.mobile.trim(),
            village:
              newFarmerForm.village.trim() || trackForm.address.trim(),
            union_name: newFarmerForm.union_name || "",
            upazila: newFarmerForm.upazila || activeTarget.upazila,
            district:
              newFarmerForm.district ||
              activeTarget.district ||
              "Mymensingh",
            latitude:
              newFarmerForm.latitude === ""
                ? null
                : Number(newFarmerForm.latitude),
            longitude:
              newFarmerForm.longitude === ""
                ? null
                : Number(newFarmerForm.longitude),
            farm_category:
              finalCategory === "Cattle" || finalCategory === "Fish"
                ? finalCategory
                : "Poultry",
            farm_type: finalCategory,
            feed_type: newFarmerForm.feed_type || "",
            remarks: newFarmerForm.remarks || trackForm.remark.trim() || "",
            status: shouldShowInVisits ? "Running" : "Active",
            last_feed_company: trackForm.feed_company,
          };
          POULTRY_STOCK_KEYS.forEach((k) => (newFarmerPayload[k] = 0));
          if (stockKey) {
            newFarmerPayload[stockKey] = effectiveBirdsForFarmAndVisit;
          }

          const { data: c1, error: cErr } = await supabase
            .from("farmers")
            .insert([newFarmerPayload])
            .select()
            .single();

          if (!cErr && c1) {
            targetFarmerId = c1.id;
          } else {
            delete newFarmerPayload.last_feed_company;
            const { data: c2, error: c2Err } = await supabase
              .from("farmers")
              .insert([newFarmerPayload])
              .select()
              .single();
            if (c2Err) throw c2Err;
            if (c2) targetFarmerId = c2.id;
          }
        }

        // 3. Save to DOC Trackings
        const insertPayload: Record<string, any> = {
          supply_id: liveStats.supplyId || activeTarget.supply_id || null,
          doc_id: normStr(activeTarget.doc_id),
          h_date: trackHDate,
          pcd: trackPcd,
          farmer_name: trackForm.farmer_name.trim(),
          address: trackForm.address.trim(),
          mobile: trackForm.mobile.trim(),
          qty: numQty,
          b_qty: numBoxes,
          feed_company: trackForm.feed_company,
          remark: trackForm.remark.trim(),
          category: finalCategory,
          breed: finalBreed,
          ag_b: trackForm.ag_b || null,
          farmer_id: targetFarmerId || null,
          created_by: currentUserId || null,
          tracked_by_name: currentUserName,
        };

        let insertedTrackingRow: any = null;
        const { data: trkInsData, error: trkInsErr } = await supabase
          .from("doc_trackings")
          .insert([insertPayload])
          .select();

        if (!trkInsErr && trkInsData && trkInsData[0]) {
          insertedTrackingRow = trkInsData[0];
        } else if (trkInsErr) {
          const withoutBQty = { ...insertPayload };
          delete withoutBQty.b_qty;
          const { data: try2Data, error: try2Err } = await supabase
            .from("doc_trackings")
            .insert([withoutBQty])
            .select();

          if (!try2Err && try2Data && try2Data[0]) {
            insertedTrackingRow = try2Data[0];
          } else {
            const fallbackPayload: Record<string, any> = {
              supply_id: liveStats.supplyId || activeTarget.supply_id || null,
              doc_id: normStr(activeTarget.doc_id),
              h_date: trackHDate,
              pcd: trackPcd,
              farmer_name: trackForm.farmer_name.trim(),
              address: trackForm.address.trim(),
              mobile: trackForm.mobile.trim(),
              qty: numQty,
              feed_company: trackForm.feed_company,
              remark: trackForm.remark.trim(),
              created_by: currentUserId || null,
            };
            const { data: retryData, error: retryErr } = await supabase
              .from("doc_trackings")
              .insert([fallbackPayload])
              .select();
            if (retryErr) throw retryErr;
            insertedTrackingRow = retryData?.[0] || fallbackPayload;
          }
        }

        const uiTrackingRow: DocTracking = {
          id:
            insertedTrackingRow?.id ||
            `local_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
          supply_id: liveStats.supplyId || activeTarget.supply_id || null,
          doc_id: normStr(activeTarget.doc_id),
          h_date: trackHDate,
          pcd: trackPcd,
          farmer_name: trackForm.farmer_name.trim(),
          address: trackForm.address.trim(),
          mobile: trackForm.mobile.trim(),
          qty: numQty,
          b_qty: numBoxes,
          feed_company: trackForm.feed_company,
          remark: trackForm.remark.trim(),
          category: finalCategory,
          breed: finalBreed,
          ag_b: trackForm.ag_b || "",
          farmer_id: targetFarmerId || null,
          created_by: currentUserId || undefined,
          tracked_by_name: currentUserName,
        };

        setTrackings((prev) => [...prev, uiTrackingRow]);

        // 4. Update Existing Farmer Stock & Create Vet Visit (Only if not "Only Track")
        let runningBatches: any[] = [];
        let existingFarmerRec: any = null;

        if (!isOnlyTrackMode && targetFarmerId) {
          existingFarmerRec =
            farmersList.find((f) => Number(f.id) === Number(targetFarmerId)) ||
            null;
          const { data: rb } = await supabase
            .from("vet_visits")
            .select("id, category, batch_label, created_at")
            .eq("farmer_id", String(targetFarmerId))
            .in("status", ["Activated", "Running"])
            .order("created_at", { ascending: true });
          runningBatches = rb || [];

          if (existingFarmerRec) {
            const prevCats = Array.from(
              new Set(runningBatches.map((r) => r.category).filter(Boolean))
            ) as string[];
            const prevForCompare =
              prevCats.length > 0
                ? prevCats
                : existingFarmerRec?.farm_type
                ? [existingFarmerRec.farm_type as string]
                : [];
            const needsChoice =
              prevForCompare.length > 0 &&
              !prevForCompare.some(
                (c) => c.toLowerCase() === finalCategory.toLowerCase()
              );
            const mode: "add" | "replace" = needsChoice ? batchMode : "add";
            const sameCatRunning = runningBatches.filter(
              (r) =>
                (r.category || "").toLowerCase() === finalCategory.toLowerCase()
            );

            const farmerUpdatePayload: Record<string, any> = {
              farm_type: finalCategory,
              farm_category:
                finalCategory === "Cattle" || finalCategory === "Fish"
                  ? finalCategory
                  : "Poultry",
              last_feed_company: trackForm.feed_company,
            };
            if (shouldShowInVisits) farmerUpdatePayload.status = "Running";

            if (stockKey) {
              const cur = Number(existingFarmerRec?.[stockKey] || 0);
              farmerUpdatePayload[stockKey] =
                sameCatRunning.length > 0
                  ? cur + effectiveBirdsForFarmAndVisit
                  : effectiveBirdsForFarmAndVisit;
              if (mode === "replace") {
                POULTRY_STOCK_KEYS.forEach((k) => {
                  if (k !== stockKey) farmerUpdatePayload[k] = 0;
                });
              }
            }

            const { error: fUpdErr } = await supabase
              .from("farmers")
              .update(farmerUpdatePayload)
              .eq("id", targetFarmerId);
            if (fUpdErr) {
              delete farmerUpdatePayload.last_feed_company;
              await supabase
                .from("farmers")
                .update(farmerUpdatePayload)
                .eq("id", targetFarmerId);
            }
          }
        }

        // 5. Add to Running Vet Visits if checked
        if (shouldShowInVisits) {
          const visitAgentId =
            existingFarmerRec?.agent_id ||
            (farmerEntryMode === "other_agent" && otherAgentId
              ? Number(otherAgentId)
              : shouldSaveFarmer && newFarmerForm.agent_id
              ? Number(newFarmerForm.agent_id)
              : feedAgentFromTarget);

          if (visitAgentId) {
            const baseFarmerName = trackForm.farmer_name.trim();
            const catLower = finalCategory.toLowerCase();
            const sameCatCount = runningBatches.filter(
              (r) => (r.category || "").toLowerCase() === catLower
            ).length;

            if (runningBatches.length > 0) {
              const counters: Record<string, number> = {};
              for (const rb of runningBatches) {
                const cat = rb.category || "Broiler";
                counters[cat] = (counters[cat] || 0) + 1;
                const lbl = String(rb.batch_label || "");
                const unmarked =
                  !lbl ||
                  lbl === baseFarmerName ||
                  (lbl.startsWith(baseFarmerName) && !lbl.includes("Batch"));
                if (unmarked) {
                  await supabase
                    .from("vet_visits")
                    .update({
                      batch_label: `${baseFarmerName} - ${cat} Batch (${counters[cat]})`,
                    })
                    .eq("id", rb.id);
                }
              }
            }

            const manualMark = trackForm.batch_mark.trim();
            const batchLabel = manualMark
              ? `${baseFarmerName} (${manualMark})`
              : runningBatches.length > 0
              ? `${baseFarmerName} - ${finalCategory} Batch (${
                  sameCatCount + 1
                })`
              : baseFarmerName;

            const nourishHatcheryName =
              companyNames.find((c) => c.toLowerCase().includes("nourish")) ||
              "Nourish";

            const visitPayload: Record<string, any> = {
              farmer_id: targetFarmerId ? String(targetFarmerId) : null,
              agent_id: String(visitAgentId),
              company_id: String(targetCompId),
              category: finalCategory,
              hatch_date: trackHDate,
              hatchery: nourishHatcheryName,
              breed: finalBreed,
              box_no: numBoxes > 0 ? String(numBoxes) : "",
              t_birds: effectiveBirdsForFarmAndVisit,
              status: "Activated",
              batch_label: batchLabel,
            };

            const { error: vErr } = await supabase
              .from("vet_visits")
              .insert([visitPayload]);

            if (vErr) {
              delete visitPayload.batch_label;
              await supabase.from("vet_visits").insert([visitPayload]);
            }
          }
        }
      }

      setIsTrackModalOpen(false);
      setEditingTracking(null);
      setFarmerEntryMode("");
      await loadAllData();
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to save tracking.");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleDeleteTracking(tr: DocTracking) {
    if (!canModifyTracking(tr)) {
      alert("You can only delete tracking entries added by you.");
      return;
    }
    if (!window.confirm("Remove this farmer tracking entry?")) return;
    try {
      const { error } = await supabase
        .from("doc_trackings")
        .delete()
        .eq("id", tr.id);
      if (error) throw error;
      setTrackings((prev) => prev.filter((item) => item.id !== tr.id));
      await loadAllData();
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to delete tracking.");
    }
  }

  const allPcdTypes = Array.from(new Set(supplies.map((s) => s.pcd))).sort();
  const categoryFilteredPcdTypes = allPcdTypes.filter((p) =>
    matchesCategoryByPcd(p, categoryFilter)
  );

  const availableAssignedUpazilas = Array.from(
    new Set(
      (assignedUpazilas.length > 0 && !isManagerOrAdmin
        ? assignedUpazilas
        : agents.map((a) => a.upazila)
      ).filter(Boolean)
    )
  ).sort();

  const allMonths = Array.from(
    new Set(supplies.map((s) => s.h_date.slice(0, 7)))
  )
    .sort()
    .reverse();

  function formatMonthLabel(ym: string) {
    const [year, month] = ym.split("-");
    const date = new Date(Number(year), Number(month) - 1, 1);
    return date.toLocaleString("en-US", { month: "short", year: "numeric" });
  }

  function handleMonthChange(ym: string) {
    setMonthFilter(ym);
    if (!ym) {
      setFromDate("");
      setToDate("");
      if (cachedPageState) {
        cachedPageState.monthFilter = "";
        cachedPageState.fromDate = "";
        cachedPageState.toDate = "";
      }
      return;
    }
    const [y, m] = ym.split("-").map(Number);
    const firstDay = `${ym}-01`;
    const lastDayNum = new Date(y, m, 0).getDate();
    const lastDay = `${ym}-${String(lastDayNum).padStart(2, "0")}`;

    setIsMultiDate(true);
    setFromDate(firstDay);
    setToDate(lastDay);
    if (cachedPageState) {
      cachedPageState.isMultiDate = true;
      cachedPageState.monthFilter = ym;
      cachedPageState.fromDate = firstDay;
      cachedPageState.toDate = lastDay;
    }
  }

  function resetAllFilters() {
    const latestDate = supplies[0]?.h_date || "";
    setFromDate(latestDate);
    setToDate(latestDate);
    setIsMultiDate(false);
    setMonthFilter("");
    setCategoryFilter("");
    setPcdFilter("");
    setUpazilaFilter("");
    setAgentDocIdFilter("");
    setOfficerFilter("");
    setTrackFilter("all");
    if (cachedPageState) {
      cachedPageState.fromDate = latestDate;
      cachedPageState.toDate = latestDate;
      cachedPageState.isMultiDate = false;
      cachedPageState.monthFilter = "";
    }
  }

  function matchesDateFilter(hDate: string) {
    const clean = normDate(hDate);
    if (monthFilter && !clean.startsWith(monthFilter)) return false;
    if (fromDate && clean < fromDate) return false;
    if (toDate && clean > toDate) return false;
    return true;
  }

  // Agents matching Upazila & Officer filter
  const baseScopedAgents = agents.filter((agent) => {
    if (upazilaFilter && agent.upazila !== upazilaFilter) return false;
    if (officerFilter) {
      const selectedOff = officers.find((o) => o.id === officerFilter);
      if (selectedOff && !selectedOff.upazilas.includes(agent.upazila)) {
        return false;
      }
    }
    return true;
  });

  // Agents that actually have DOC on the selected dates, Category & DOC type
  const agentsWithDocInScope = baseScopedAgents.filter((agent) =>
    supplies.some(
      (s) =>
        normStr(s.doc_id) === normStr(agent.doc_id) &&
        matchesDateFilter(s.h_date) &&
        matchesCategoryByPcd(s.pcd, categoryFilter) &&
        (!pcdFilter || normCode(s.pcd) === normCode(pcdFilter))
    )
  );

  const matchingAgents = baseScopedAgents.filter((agent) => {
    if (
      agentDocIdFilter &&
      normStr(agent.doc_id) !== normStr(agentDocIdFilter)
    ) {
      return false;
    }
    return true;
  });

  const matchingAgentDocIds = new Set(
    matchingAgents.map((a) => normStr(a.doc_id))
  );

  const filteredSupplies = supplies.filter((s) => {
    if (!matchingAgentDocIds.has(normStr(s.doc_id))) return false;
    if (!matchesCategoryByPcd(s.pcd, categoryFilter)) return false;
    const matchPcd = !pcdFilter || normCode(s.pcd) === normCode(pcdFilter);
    return matchesDateFilter(s.h_date) && matchPcd;
  });

  const filteredTrackings = trackings.filter((t) => {
    if (!matchingAgentDocIds.has(normStr(t.doc_id))) return false;
    if (!matchesCategoryByPcd(t.pcd, categoryFilter)) return false;
    const matchPcd = !pcdFilter || normCode(t.pcd) === normCode(pcdFilter);
    return matchesDateFilter(t.h_date) && matchPcd;
  });

  const visibleAgents = matchingAgents.filter((agent) =>
    filteredSupplies.some((s) => normStr(s.doc_id) === normStr(agent.doc_id))
  );

  const dateRangeLabel =
    fromDate && toDate
      ? fromDate === toDate
        ? fromDate
        : `${fromDate} to ${toDate}`
      : fromDate
      ? `From ${fromDate}`
      : toDate
      ? `Up to ${toDate}`
      : monthFilter
      ? formatMonthLabel(monthFilter)
      : "All Dates";

  const totalSupplyQty = filteredSupplies.reduce(
    (sum, s) => sum + (Number(s.qty) || 0),
    0
  );
  const totalTrackedQty = filteredTrackings.reduce(
    (sum, t) => sum + (Number(t.qty) || 0),
    0
  );
  const totalUntrackedQty = Math.max(0, totalSupplyQty - totalTrackedQty);

  const trackedPercent =
    totalSupplyQty > 0
      ? ((totalTrackedQty / totalSupplyQty) * 100).toFixed(1)
      : "0.0";
  const untrackedPercent =
    totalSupplyQty > 0
      ? ((totalUntrackedQty / totalSupplyQty) * 100).toFixed(1)
      : "0.0";

  function openExportModal() {
    setExportSelectedTypes(
      pcdFilter ? [pcdFilter] : [...categoryFilteredPcdTypes]
    );
    setIsExportModalOpen(true);
  }

  function toggleExportDocType(type: string) {
    setExportSelectedTypes((prev) =>
      prev.includes(type) ? prev.filter((t) => t !== type) : [...prev, type]
    );
  }

  async function handleExportExcel() {
    if (exportSelectedTypes.length === 0) {
      alert("Please select at least one DOC Type to export.");
      return;
    }

    const exportRows: Record<string, any>[] = [];
    let grandTotalSupply = 0;
    let grandTotalTracked = 0;

    visibleAgents.forEach((agent) => {
      const agentSupplies = filteredSupplies.filter(
        (s) =>
          normStr(s.doc_id) === normStr(agent.doc_id) &&
          exportSelectedTypes.includes(normCode(s.pcd))
      );

      const groups = new Map<
        string,
        { h_date: string; pcd: string; qty: number; b_qty: number }
      >();
      agentSupplies.forEach((s) => {
        const key = `${normDate(s.h_date)}_${normCode(s.pcd)}`;
        const prev = groups.get(key);
        if (prev) {
          prev.qty += Number(s.qty) || 0;
          prev.b_qty += Number(s.b_qty) || 0;
        } else {
          groups.set(key, {
            h_date: normDate(s.h_date),
            pcd: normCode(s.pcd),
            qty: Number(s.qty) || 0,
            b_qty: Number(s.b_qty) || 0,
          });
        }
      });

      groups.forEach((g) => {
        grandTotalSupply += g.qty;
        const docPerBox = g.qty > 0 && g.b_qty > 0 ? g.qty / g.b_qty : 0;

        const matchingTracks = filteredTrackings.filter(
          (t) =>
            normStr(t.doc_id) === normStr(agent.doc_id) &&
            normDate(t.h_date) === g.h_date &&
            normCode(t.pcd) === g.pcd
        );

        if (matchingTracks.length === 0) {
          exportRows.push({
            "H.Date": g.h_date,
            "DOC.ID": agent.doc_id,
            Agent: agent.agent_name,
            "Agent Address": `${agent.upazila}, ${agent.district}`,
            "DOC Qty": g.qty,
            "DOC Type": g.pcd,
            "AG/B+": "-",
            Farmer: "-",
            "Farmer Address": "-",
            "Farmer Mobile": "-",
            "Tracked Qty": 0,
            "Box Qty": 0,
            "Feed Company": "-",
            "Tracked By": "-",
            Remark: "Untracked",
          });
        } else {
          matchingTracks.forEach((t, index) => {
            grandTotalTracked += Number(t.qty) || 0;
            const trBox =
              t.b_qty !== undefined && t.b_qty !== null
                ? Number(t.b_qty)
                : docPerBox > 0
                ? Math.round(Number(t.qty || 0) / docPerBox)
                : 0;

            exportRows.push({
              "H.Date": g.h_date,
              "DOC.ID": agent.doc_id,
              Agent: agent.agent_name,
              "Agent Address": `${agent.upazila}, ${agent.district}`,
              "DOC Qty": index === 0 ? g.qty : "",
              "DOC Type": g.pcd,
              "AG/B+": t.ag_b || "-",
              Farmer: t.farmer_name,
              "Farmer Address": t.address || "-",
              "Farmer Mobile": t.mobile || "-",
              "Tracked Qty": Number(t.qty) || 0,
              "Box Qty": trBox,
              "Feed Company": t.feed_company,
              "Tracked By": t.tracked_by_name || "-",
              Remark: t.remark || "-",
            });
          });
        }
      });
    });

    if (exportRows.length === 0) {
      alert("No tracking data found for the selected DOC Type(s).");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("DOC_Tracking");
    worksheet.views = [{ state: "frozen", ySplit: 1 }];

    const headers = Object.keys(exportRows[0]);
    const headerRow = worksheet.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "1E3A8A" },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
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
    headerRow.height = 26;

    exportRows.forEach((r, idx) => {
      const rowValues = headers.map((h) => r[h]);
      const row = worksheet.addRow(rowValues);
      const rowBg = idx % 2 === 0 ? "FFFFFF" : "F8FAFC";

      row.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: rowBg },
        };
        cell.font = { color: { argb: "0F172A" } };
        cell.alignment = { vertical: "middle", wrapText: true };
        cell.border = {
          top: { style: "thin", color: { argb: "E2E8F0" } },
          left: { style: "thin", color: { argb: "E2E8F0" } },
          bottom: { style: "thin", color: { argb: "E2E8F0" } },
          right: { style: "thin", color: { argb: "E2E8F0" } },
        };
      });
      row.height = 22;
    });

    const totalRowValues = headers.map((h) => {
      if (h === "H.Date") return "TOTAL";
      if (h === "DOC Qty") return grandTotalSupply;
      if (h === "Tracked Qty") return grandTotalTracked;
      if (h === "Remark")
        return `Untracked: ${Math.max(0, grandTotalSupply - grandTotalTracked)}`;
      return "";
    });

    const totalRow = worksheet.addRow(totalRowValues);
    totalRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FEF3C7" },
      };
      cell.font = { bold: true, color: { argb: "92400E" } };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = {
        top: { style: "medium", color: { argb: "D97706" } },
        left: { style: "thin", color: { argb: "CBD5E1" } },
        bottom: { style: "medium", color: { argb: "D97706" } },
        right: { style: "thin", color: { argb: "CBD5E1" } },
      };
    });
    totalRow.height = 25;

    worksheet.columns.forEach((column, i) => {
      let maxLen = headers[i].length;
      exportRows.forEach((row) => {
        const val =
          row[headers[i]] !== undefined && row[headers[i]] !== null
            ? String(row[headers[i]])
            : "";
        if (val.length > maxLen) maxLen = val.length;
      });
      column.width = Math.max(maxLen + 5, 12);
    });

    worksheet.autoFilter = {
      from: "A1",
      to: `${String.fromCharCode(64 + headers.length)}1`,
    };

    const fileName = `DOC_Tracking_${dateRangeLabel || "Report"}_${exportSelectedTypes.join("-")}.xlsx`;

    try {
      const excelBuffer = await workbook.xlsx.writeBuffer();

      if (Capacitor.isNativePlatform()) {
        let binary = "";
        const bytes = new Uint8Array(excelBuffer);
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
          } catch {}
        }
      } else {
        const blob = new Blob([excelBuffer], {
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
      console.error("Export error:", err);
      alert("Failed to export: " + (err?.message || ""));
    } finally {
      setIsExportModalOpen(false);
    }
  }

  const reportSupplies = supplies.filter((s) => {
    if (!matchingAgentDocIds.has(normStr(s.doc_id))) return false;
    const matchFrom = !reportFromDate || s.h_date >= reportFromDate;
    const matchTo = !reportToDate || s.h_date <= reportToDate;
    const matchPcd =
      !reportPcdFilter || normCode(s.pcd) === normCode(reportPcdFilter);
    return matchFrom && matchTo && matchPcd;
  });

  const reportTrackings = trackings.filter((t) => {
    if (!matchingAgentDocIds.has(normStr(t.doc_id))) return false;
    const matchFrom = !reportFromDate || t.h_date >= reportFromDate;
    const matchTo = !reportToDate || t.h_date <= reportToDate;
    const matchPcd =
      !reportPcdFilter || normCode(t.pcd) === normCode(reportPcdFilter);
    return matchFrom && matchTo && matchPcd;
  });

  const repTotalDoc = reportSupplies.reduce(
    (sum, s) => sum + (Number(s.qty) || 0),
    0
  );
  const repTotalTracked = reportTrackings.reduce(
    (sum, t) => sum + (Number(t.qty) || 0),
    0
  );
  const repTotalUntracked = Math.max(0, repTotalDoc - repTotalTracked);
  const repTrackedPct =
    repTotalDoc > 0
      ? ((repTotalTracked / repTotalDoc) * 100).toFixed(1)
      : "0.0";
  const repUntrackedPct =
    repTotalDoc > 0
      ? ((repTotalUntracked / repTotalDoc) * 100).toFixed(1)
      : "0.0";

  const allReportCompanies = Array.from(
    new Set([
      ...companyNames,
      ...reportTrackings.map((t) => t.feed_company).filter(Boolean),
    ])
  );

  const companyBreakdown = allReportCompanies.map((comp) => {
    const compQty = reportTrackings
      .filter((t) => t.feed_company?.toLowerCase() === comp.toLowerCase())
      .reduce((sum, t) => sum + (Number(t.qty) || 0), 0);

    const pctOfTracked =
      repTotalTracked > 0
        ? ((compQty / repTotalTracked) * 100).toFixed(1)
        : "0.0";
    const pctOfTotal =
      repTotalDoc > 0 ? ((compQty / repTotalDoc) * 100).toFixed(1) : "0.0";

    return { company: comp, qty: compQty, pctOfTracked, pctOfTotal };
  });

  const pcdBreakdown = allPcdTypes.map((pcd) => {
    const supQty = supplies
      .filter((s) => {
        if (!matchingAgentDocIds.has(normStr(s.doc_id))) return false;
        const matchFrom = !reportFromDate || s.h_date >= reportFromDate;
        const matchTo = !reportToDate || s.h_date <= reportToDate;
        return matchFrom && matchTo && normCode(s.pcd) === normCode(pcd);
      })
      .reduce((sum, s) => sum + (Number(s.qty) || 0), 0);

    const trkQty = trackings
      .filter((t) => {
        if (!matchingAgentDocIds.has(normStr(t.doc_id))) return false;
        const matchFrom = !reportFromDate || t.h_date >= reportFromDate;
        const matchTo = !reportToDate || t.h_date <= reportToDate;
        return matchFrom && matchTo && normCode(t.pcd) === normCode(pcd);
      })
      .reduce((sum, t) => sum + (Number(t.qty) || 0), 0);

    const untrkQty = Math.max(0, supQty - trkQty);
    const pct = supQty > 0 ? ((trkQty / supQty) * 100).toFixed(1) : "0.0";

    return { pcd, supQty, trkQty, untrkQty, pct };
  });

  // Build unified blocks per agent (latest h_date first)
  function buildAgentBlocks(agent: DocAgent) {
    const cleanAgentDocId = normStr(agent.doc_id);
    const sup = filteredSupplies.filter(
      (x) => normStr(x.doc_id) === cleanAgentDocId
    );
    const keys = new Map<string, { h_date: string; pcd: string }>();
    sup.forEach((x) => {
      const d = normDate(x.h_date);
      const p = normCode(x.pcd);
      keys.set(`${d}__${p}`, { h_date: d, pcd: p });
    });

    return Array.from(keys.values())
      .sort(
        (a, b) =>
          b.h_date.localeCompare(a.h_date) || a.pcd.localeCompare(b.pcd)
      )
      .map(({ h_date, pcd }) => {
        const rows = sup.filter(
          (x) => normDate(x.h_date) === h_date && normCode(x.pcd) === pcd
        );
        const total = rows.reduce((sum, r) => sum + (Number(r.qty) || 0), 0);
        const bTotal = rows.reduce((sum, r) => sum + (Number(r.b_qty) || 0), 0);
        const docPerBox = total > 0 && bTotal > 0 ? total / bTotal : 0;
        const agBText =
          Array.from(new Set(rows.map((r) => r.ag_b).filter(Boolean))).join(
            ", "
          ) || "-";
        const trk = filteredTrackings.filter(
          (t) =>
            normStr(t.doc_id) === cleanAgentDocId &&
            normDate(t.h_date) === h_date &&
            normCode(t.pcd) === pcd
        );
        const tracked = trk.reduce((sum, r) => sum + (Number(r.qty) || 0), 0);
        return {
          h_date,
          pcd,
          rows,
          total,
          bTotal,
          docPerBox,
          agBText,
          trk,
          tracked,
          untracked: Math.max(0, total - tracked),
        };
      });
  }

  const agentBlocksList = visibleAgents
    .map((agent) => {
      const allBlocks = buildAgentBlocks(agent);
      const filteredBlocks = allBlocks.filter(
        (b) =>
          trackFilter === "all" ||
          (trackFilter === "tracked" ? b.tracked > 0 : b.untracked > 0)
      );
      return { agent, blocks: filteredBlocks, allBlocks };
    })
    .filter((x) => x.blocks.length > 0);

  // Modal dynamic options for H.Date and DOC Type
  function modalHatchDatesForAgent(): string[] {
    if (!activeTarget) return [];
    const cleanId = normStr(activeTarget.doc_id);
    return Array.from(
      new Set(
        supplies
          .filter(
            (x) =>
              normStr(x.doc_id) === cleanId &&
              matchesDateFilter(x.h_date) &&
              matchesCategoryByPcd(x.pcd, categoryFilter) &&
              (!pcdFilter || normCode(x.pcd) === normCode(pcdFilter))
          )
          .map((x) => normDate(x.h_date))
      )
    ).sort((a, b) => b.localeCompare(a));
  }

  function modalDocTypesForDate(hDate: string): string[] {
    if (!activeTarget) return [];
    const cleanId = normStr(activeTarget.doc_id);
    const cleanDate = normDate(hDate);
    return Array.from(
      new Set(
        supplies
          .filter(
            (x) =>
              normStr(x.doc_id) === cleanId &&
              normDate(x.h_date) === cleanDate &&
              matchesCategoryByPcd(x.pcd, categoryFilter) &&
              (!pcdFilter || normCode(x.pcd) === normCode(pcdFilter))
          )
          .map((x) => normCode(x.pcd))
      )
    );
  }

  function handleModalDateChange(newDate: string) {
    if (!activeTarget) return;
    const typesForDate = modalDocTypesForDate(newDate);
    const nextPcd = typesForDate.includes(normCode(trackForm.pcd))
      ? normCode(trackForm.pcd)
      : typesForDate[0] || activeTarget.pcd;

    const stats = getStatsForAgentDatePcd(
      activeTarget.doc_id,
      newDate,
      nextPcd,
      editingTracking?.id
    );
    const mapped = mapPcdToCategoryAndBreed(nextPcd);
    const autoBoxes = computeAutoBoxQty(stats.remaining, stats.docPerBox);

    setActiveTarget({
      ...activeTarget,
      h_date: newDate,
      pcd: nextPcd,
      supply_id: stats.supplyId,
      remainingQty: stats.remaining,
    });
    setTrackForm((prev) => ({
      ...prev,
      h_date: newDate,
      pcd: nextPcd,
      category: mapped.category,
      breed: mapped.breed,
      qty: String(stats.remaining),
      box_qty: autoBoxes,
      ag_b: stats.agBOptions.length === 1 ? stats.agBOptions[0] : "",
    }));
  }

  function handleModalPcdChange(newPcd: string) {
    if (!activeTarget) return;
    const curDate = normDate(trackForm.h_date || activeTarget.h_date);
    const stats = getStatsForAgentDatePcd(
      activeTarget.doc_id,
      curDate,
      newPcd,
      editingTracking?.id
    );
    const mapped = mapPcdToCategoryAndBreed(newPcd);
    const autoBoxes = computeAutoBoxQty(stats.remaining, stats.docPerBox);

    setActiveTarget({
      ...activeTarget,
      pcd: newPcd,
      supply_id: stats.supplyId,
      remainingQty: stats.remaining,
    });
    setTrackForm((prev) => ({
      ...prev,
      pcd: newPcd,
      category: mapped.category,
      breed: mapped.breed,
      qty: String(stats.remaining),
      box_qty: autoBoxes,
      ag_b: stats.agBOptions.length === 1 ? stats.agBOptions[0] : "",
    }));
  }

  const matchedFarmerId: number | null = (() => {
    if (!isTrackModalOpen || editingTracking || !activeTarget) return null;
    if (farmerEntryMode === "only_track") return null;
    if (selectedFarmerId) return Number(selectedFarmerId);
    const nm = trackForm.farmer_name.trim().toLowerCase();
    if (!nm) return null;
    const fa = getCollaboratedFeedAgentId(activeTarget);
    const f = farmersList.find(
      (x) =>
        x.farmer_name?.trim().toLowerCase() === nm &&
        (!fa || Number(x.agent_id) === Number(fa))
    );
    return f ? f.id : null;
  })();

  // Smart Duplicate Prevention: Background similarity search on Farmer Name & Mobile
  const similarFarmersList = useMemo(() => {
    if (
      editingTracking ||
      (farmerEntryMode !== "create_farmer" && farmerEntryMode !== "only_track")
    ) {
      return [];
    }

    const rawInput = trackForm.farmer_name.trim();
    const rawMobile = trackForm.mobile.trim();
    if (rawInput.length < 2 && rawMobile.length < 4) return [];

    const normInput = normalizeFarmerStr(rawInput);
    const inputWords = rawInput
      .toLowerCase()
      .split(/\s+/)
      .filter((w) => w.length >= 2);

    return farmersList
      .map((f) => {
        let score = 0;
        const fName = (f.farmer_name || "").toLowerCase();
        const normFName = normalizeFarmerStr(f.farmer_name);
        const fMobile = (f.mobile || "").trim();

        if (fName === rawInput.toLowerCase()) score += 100;
        else if (normInput.length >= 3 && normFName === normInput) score += 90;
        else if (fName.includes(rawInput.toLowerCase())) score += 75;
        else if (normInput.length >= 3 && normFName.includes(normInput))
          score += 65;

        inputWords.forEach((w) => {
          if (fName.includes(w)) score += 25;
        });

        if (rawMobile.length >= 4 && fMobile.includes(rawMobile)) {
          score += 80;
        }

        const agentObj = feedAgents.find(
          (a) => Number(a.id) === Number(f.agent_id)
        );

        return {
          farmer: f,
          agentName: agentObj?.agent_name || "-",
          companyName: f.companies?.company_name || "-",
          score,
        };
      })
      .filter((item) => item.score >= 25)
      .sort((a, b) => b.score - a.score)
      .slice(0, 6);
  }, [
    trackForm.farmer_name,
    trackForm.mobile,
    farmersList,
    feedAgents,
    farmerEntryMode,
    editingTracking,
  ]);

  useEffect(() => {
    let cancelled = false;
    async function detect() {
      if (!matchedFarmerId) {
        setBatchCtx(null);
        return;
      }
      const { data } = await supabase
        .from("vet_visits")
        .select("id, category, batch_label, created_at")
        .eq("farmer_id", String(matchedFarmerId))
        .in("status", ["Activated", "Running"])
        .order("created_at", { ascending: true });
      if (cancelled) return;
      const running = data || [];
      let prevCats = Array.from(
        new Set(running.map((r: any) => r.category).filter(Boolean))
      ) as string[];
      const inactive = running.length === 0;
      if (inactive) {
        const fr: any = farmersList.find(
          (f) => Number(f.id) === matchedFarmerId
        );
        if (fr?.farm_type) {
          prevCats = [fr.farm_type];
        } else if (fr) {
          const found = Object.entries(STOCK_KEY_BY_CATEGORY).find(
            ([, key]) => Number(fr[key] || 0) > 0
          );
          if (found) prevCats = [found[0]];
        }
      }
      setBatchCtx({ farmerId: matchedFarmerId, running, prevCats, inactive });
    }
    detect();
    return () => {
      cancelled = true;
    };
  }, [matchedFarmerId, isTrackModalOpen]);

  const needsBatchChoice =
    !!batchCtx &&
    batchCtx.prevCats.length > 0 &&
    !batchCtx.prevCats.some(
      (c) => c.toLowerCase() === trackForm.category.toLowerCase()
    );

  useEffect(() => {
    setBatchMode("add");
  }, [trackForm.category, matchedFarmerId]);

  useEffect(() => {
    if (
      (!saveToFarmers &&
        farmerEntryMode !== "create_farmer" &&
        farmerEntryMode !== "other_agent") ||
      geoDistricts.length > 0
    )
      return;
    supabase
      .from("districts")
      .select("*")
      .order("name")
      .then(({ data }) => setGeoDistricts(data || []));
  }, [saveToFarmers, farmerEntryMode]);

  useEffect(() => {
    const d = geoDistricts.find((x) => x.name === newFarmerForm.district);
    if (!d) {
      setGeoUpazilas([]);
      return;
    }
    supabase
      .from("upazilas")
      .select("*")
      .eq("district_id", d.id)
      .order("name")
      .then(({ data }) => setGeoUpazilas(data || []));
  }, [newFarmerForm.district, geoDistricts]);

  useEffect(() => {
    const u = geoUpazilas.find((x) => x.name === newFarmerForm.upazila);
    if (!u) {
      setGeoUnions([]);
      return;
    }
    supabase
      .from("unions")
      .select("*")
      .eq("upazila_id", u.id)
      .order("name")
      .then(({ data }) => setGeoUnions(data || []));
  }, [newFarmerForm.upazila, geoUpazilas]);

  function initNewFarmerFormDefaults() {
    if (!activeTarget) return;
    const comp = companiesList.find(
      (c) =>
        c.company_name?.trim().toLowerCase() ===
        trackForm.feed_company.trim().toLowerCase()
    );
    const nourish = companiesList.find((c) =>
      c.company_name?.toLowerCase().includes("nourish")
    );
    const feedAgentId = getCollaboratedFeedAgentId(activeTarget);
    setNewFarmerForm({
      company_id: String((comp || nourish || companiesList[0])?.id || ""),
      district: activeTarget.district || "Mymensingh",
      upazila: activeTarget.upazila || "",
      agent_id: feedAgentId ? String(feedAgentId) : "",
      union_name: "",
      village: trackForm.address,
      latitude: "",
      longitude: "",
      feed_type: "",
      remarks: "",
    });
  }

  const nfAgentOptions = feedAgents.filter((a) => {
    if (
      newFarmerForm.company_id &&
      a.company_id &&
      String(a.company_id) !== String(newFarmerForm.company_id)
    )
      return false;
    if (
      newFarmerForm.upazila &&
      a.upazila &&
      a.upazila !== newFarmerForm.upazila
    )
      return false;
    return true;
  });

  // Other Agents List (Excludes current agent)
  const currentCollaboratedAgentId = getCollaboratedFeedAgentId(activeTarget);
  const otherAgentsOptions = feedAgents.filter((a) => {
    if (
      currentCollaboratedAgentId &&
      Number(a.id) === Number(currentCollaboratedAgentId)
    ) {
      return false;
    }
    if (
      otherCompanyId &&
      a.company_id &&
      String(a.company_id) !== String(otherCompanyId)
    ) {
      return false;
    }
    if (otherUpazila && a.upazila && a.upazila !== otherUpazila) {
      return false;
    }
    return true;
  });

  const otherAgentFarmersList = farmersList.filter(
    (f) => otherAgentId && String(f.agent_id) === String(otherAgentId)
  );

  const allFeedUpazilas = Array.from(
    new Set(feedAgents.map((a) => a.upazila).filter(Boolean))
  ).sort();

  function agBadge(v?: string) {
    if (!v) return null;
    const isB = v.trim().toUpperCase().startsWith("B");
    return (
      <span
        className={`px-1 py-0 rounded text-[9px] font-bold ${
          isB ? "bg-purple-100 text-purple-700" : "bg-sky-100 text-sky-700"
        }`}
      >
        {v}
      </span>
    );
  }

  const collaboratedFarmers = getCollaboratedFarmersForTarget(activeTarget);

  // Live modal stats for selected H.Date & DOC Type
  const modalSelectedStats = activeTarget
    ? getStatsForAgentDatePcd(
        activeTarget.doc_id,
        trackForm.h_date || activeTarget.h_date,
        trackForm.pcd || activeTarget.pcd,
        editingTracking?.id
      )
    : {
        total: 0,
        bTotal: 0,
        docPerBox: 0,
        tracked: 0,
        remaining: 0,
        agBOptions: ["AG", "B+"],
      };

  const selectedNewFarmerUpazilaObj = geoUpazilas.find(
    (u) => u.name === newFarmerForm.upazila
  );
  const selectedNewFarmerDistrictObj = geoDistricts.find(
    (d) => d.name === newFarmerForm.district
  );

  const isTrackingFormVisible =
    !!editingTracking ||
    (farmerEntryMode === "this_agent" && !!selectedFarmerId) ||
    (farmerEntryMode === "other_agent" && !!selectedFarmerId) ||
    farmerEntryMode === "create_farmer" ||
    farmerEntryMode === "only_track";

  return (
    <div className="p-2 sm:p-4 max-w-7xl mx-auto space-y-2.5">
      {/* HEADER BAR */}
      <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-2 bg-white border border-slate-200 rounded-xl px-3 py-2 shadow-sm">
        <div className="flex flex-wrap items-center gap-1.5">
          <span className="px-2 py-0.5 rounded-lg bg-slate-100 border border-slate-200 text-slate-700 text-[11px] sm:text-xs font-medium">
            Date: <strong className="text-slate-900">{dateRangeLabel}</strong>
          </span>

          <button
            type="button"
            onClick={() => setTrackFilter("all")}
            className={`px-2 py-0.5 rounded-lg border text-[11px] sm:text-xs font-medium transition ${
              trackFilter === "all"
                ? "bg-slate-800 border-slate-800 text-white"
                : "bg-slate-100 border-slate-200 text-slate-700 hover:bg-slate-200"
            }`}
          >
            Total DOC: <strong>{totalSupplyQty.toLocaleString()}</strong>
          </button>

          <button
            type="button"
            onClick={() =>
              setTrackFilter(trackFilter === "tracked" ? "all" : "tracked")
            }
            className={`px-2 py-0.5 rounded-lg border text-[11px] sm:text-xs font-medium transition ${
              trackFilter === "tracked"
                ? "bg-emerald-600 border-emerald-700 text-white ring-2 ring-emerald-200"
                : "bg-emerald-50 border-emerald-200 text-emerald-700 hover:bg-emerald-100"
            }`}
          >
            Tracked:{" "}
            <strong>
              {totalTrackedQty.toLocaleString()} ({trackedPercent}%)
            </strong>
          </button>

          <button
            type="button"
            onClick={() =>
              setTrackFilter(trackFilter === "untracked" ? "all" : "untracked")
            }
            className={`px-2 py-0.5 rounded-lg border text-[11px] sm:text-xs font-medium transition ${
              trackFilter === "untracked"
                ? "bg-amber-500 border-amber-600 text-white ring-2 ring-amber-200"
                : "bg-amber-50 border-amber-200 text-amber-700 hover:bg-amber-100"
            }`}
          >
            Untracked:{" "}
            <strong>
              {totalUntrackedQty.toLocaleString()} ({untrackedPercent}%)
            </strong>
          </button>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            onClick={() => initTrackingPage(true)}
            title="Refresh Data"
            className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs transition"
          >
            <FaSyncAlt className={loading ? "animate-spin" : ""} />
          </button>

          {isManagerOrAdmin && (
            <>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleExcelUpload}
                className="hidden"
              />
              <button
                onClick={() => fileInputRef.current?.click()}
                disabled={importing}
                className="inline-flex items-center gap-1 bg-slate-700 hover:bg-slate-800 text-white px-2.5 py-1.5 rounded-lg text-[11px] sm:text-xs font-medium transition"
              >
                {importing ? (
                  <FaSpinner className="animate-spin" />
                ) : (
                  <FaFileUpload />
                )}
                Import
              </button>
            </>
          )}

          <button
            onClick={openExportModal}
            className="inline-flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1.5 rounded-lg text-[11px] sm:text-xs font-medium transition"
          >
            <FaFileExcel />
            Export
          </button>

          <button
            onClick={() => setIsReportOpen(true)}
            className="inline-flex items-center gap-1 bg-amber-500 hover:bg-amber-600 text-white px-2.5 py-1.5 rounded-lg text-[11px] sm:text-xs font-medium transition"
          >
            <FaChartPie />
            Report
          </button>
        </div>
      </div>

      {/* COMPACT FILTER BAR (INCLUDING CATEGORY FILTER: BROILER / LAYER / SONALI) */}
      <div className="bg-white border border-slate-200 rounded-xl p-2 shadow-sm">
        <div
          className={`grid grid-cols-2 sm:grid-cols-4 ${
            isMultiDate
              ? isManagerOrAdmin
                ? "lg:grid-cols-8"
                : "lg:grid-cols-7"
              : isManagerOrAdmin
              ? "lg:grid-cols-7"
              : "lg:grid-cols-6"
          } gap-1.5 items-center`}
        >
          {/* Single Date OR From-To Date Selector with small Multi Checkbox */}
          <div
            className={`${
              isMultiDate
                ? "col-span-2 sm:col-span-4 lg:col-span-2"
                : "col-span-2 sm:col-span-2 lg:col-span-1"
            } flex items-center gap-1.5 min-w-0`}
          >
            {!isMultiDate ? (
              <div className="flex-1 min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                <span className="text-[10px] font-semibold text-slate-400 mr-1.5 shrink-0">
                  Date:
                </span>
                <input
                  type="date"
                  value={fromDate}
                  onChange={(e) => {
                    const val = e.target.value;
                    setFromDate(val);
                    setToDate(val);
                    setMonthFilter("");
                    if (cachedPageState) {
                      cachedPageState.fromDate = val;
                      cachedPageState.toDate = val;
                      cachedPageState.monthFilter = "";
                    }
                  }}
                  className="w-full min-w-0 text-[11px] sm:text-xs outline-none bg-transparent text-slate-700"
                />
              </div>
            ) : (
              <div className="flex-1 min-w-0 grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-semibold text-slate-400 mr-1.5 shrink-0">
                    From:
                  </span>
                  <input
                    type="date"
                    value={fromDate}
                    onChange={(e) => {
                      setFromDate(e.target.value);
                      setMonthFilter("");
                      if (cachedPageState) {
                        cachedPageState.fromDate = e.target.value;
                        cachedPageState.monthFilter = "";
                      }
                    }}
                    className="w-full min-w-0 text-[11px] sm:text-xs outline-none bg-transparent text-slate-700"
                  />
                </div>
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-semibold text-slate-400 mr-1.5 shrink-0">
                    To:
                  </span>
                  <input
                    type="date"
                    value={toDate}
                    onChange={(e) => {
                      setToDate(e.target.value);
                      setMonthFilter("");
                      if (cachedPageState) {
                        cachedPageState.toDate = e.target.value;
                        cachedPageState.monthFilter = "";
                      }
                    }}
                    className="w-full min-w-0 text-[11px] sm:text-xs outline-none bg-transparent text-slate-700"
                  />
                </div>
              </div>
            )}

            <label
              title="Select multiple dates (From - To)"
              className="shrink-0 inline-flex items-center gap-1 px-2 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 text-[10px] sm:text-[11px] font-semibold text-slate-700 cursor-pointer select-none"
            >
              <input
                type="checkbox"
                checked={isMultiDate}
                onChange={(e) => {
                  const checked = e.target.checked;
                  setIsMultiDate(checked);
                  if (!checked) {
                    const single =
                      fromDate || toDate || supplies[0]?.h_date || "";
                    setFromDate(single);
                    setToDate(single);
                    setMonthFilter("");
                    if (cachedPageState) {
                      cachedPageState.isMultiDate = false;
                      cachedPageState.fromDate = single;
                      cachedPageState.toDate = single;
                      cachedPageState.monthFilter = "";
                    }
                  } else if (cachedPageState) {
                    cachedPageState.isMultiDate = true;
                  }
                }}
                className="w-3.5 h-3.5 accent-emerald-600"
              />
              <span>Multi</span>
            </label>
          </div>

          <select
            value={monthFilter}
            onChange={(e) => handleMonthChange(e.target.value)}
            className="border border-slate-300 rounded-lg px-2 py-1.5 text-[11px] sm:text-xs bg-white outline-none"
          >
            <option value="">All Months</option>
            {allMonths.map((ym) => (
              <option key={ym} value={ym}>
                {formatMonthLabel(ym)}
              </option>
            ))}
          </select>

          {/* NEW CATEGORY FILTER: Broiler / Layer / Sonali */}
          <select
            value={categoryFilter}
            onChange={(e) => {
              setCategoryFilter(e.target.value);
              setPcdFilter("");
            }}
            className="border border-slate-300 rounded-lg px-2 py-1.5 text-[11px] sm:text-xs bg-white outline-none font-semibold text-slate-700"
          >
            <option value="">All Categories</option>
            <option value="Broiler">🐔 Broiler (AA, EP, RO)</option>
            <option value="Layer">🥚 Layer (SH, BO)</option>
            <option value="Sonali">🐓 Sonali (CB, HP, CW, BR, BP)</option>
          </select>

          <select
            value={pcdFilter}
            onChange={(e) => setPcdFilter(e.target.value)}
            className="border border-slate-300 rounded-lg px-2 py-1.5 text-[11px] sm:text-xs bg-white outline-none"
          >
            <option value="">All DOC Types</option>
            {categoryFilteredPcdTypes.map((pcd) => (
              <option key={pcd} value={pcd}>
                {pcd} ({mapPcdToCategoryAndBreed(pcd).breed})
              </option>
            ))}
          </select>

          {/* Assigned Upazila Dropdown */}
          <select
            value={upazilaFilter}
            onChange={(e) => {
              setUpazilaFilter(e.target.value);
              setAgentDocIdFilter("");
            }}
            className="border border-slate-300 rounded-lg px-2 py-1.5 text-[11px] sm:text-xs bg-white outline-none"
          >
            <option value="">All Upazilas</option>
            {availableAssignedUpazilas.map((u) => (
              <option key={u} value={u}>
                {u}
              </option>
            ))}
          </select>

          {isManagerOrAdmin && (
            <select
              value={officerFilter}
              onChange={(e) => {
                setOfficerFilter(e.target.value);
                setAgentDocIdFilter("");
              }}
              className="border border-slate-300 rounded-lg px-2 py-1.5 text-[11px] sm:text-xs bg-white outline-none"
            >
              <option value="">All Officers</option>
              {officers
                .filter((o) => o.upazilas.length > 0)
                .map((off) => (
                  <option key={off.id} value={off.id}>
                    {off.full_name}
                  </option>
                ))}
            </select>
          )}

          {/* Dropdown of Agents that have DOC in the filtered dates + Reset */}
          <div className="col-span-2 sm:col-span-1 flex items-center gap-1">
            <select
              value={agentDocIdFilter}
              onChange={(e) => setAgentDocIdFilter(e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-2 py-1.5 text-[11px] sm:text-xs bg-white outline-none"
            >
              <option value="">
                All DOC Agents ({agentsWithDocInScope.length})
              </option>
              {agentsWithDocInScope.map((ag) => (
                <option key={ag.doc_id} value={ag.doc_id}>
                  {ag.doc_id} - {ag.agent_name}
                </option>
              ))}
            </select>

            <button
              onClick={resetAllFilters}
              title="Reset Filters"
              className="shrink-0 inline-flex items-center justify-center rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2 py-1.5 text-xs font-semibold text-slate-700 transition"
            >
              <RotateCcw size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* ALERTS */}
      {errorMsg && (
        <div className="flex items-center justify-between bg-red-50 border border-red-200 text-red-700 px-3 py-2 rounded-xl text-xs">
          <span>{errorMsg}</span>
          <button onClick={() => setErrorMsg("")}>
            <FaTimes />
          </button>
        </div>
      )}

      {successMsg && (
        <div className="flex items-center justify-between bg-emerald-50 border border-emerald-200 text-emerald-700 px-3 py-2 rounded-xl text-xs">
          <span>{successMsg}</span>
          <button onClick={() => setSuccessMsg("")}>
            <FaTimes />
          </button>
        </div>
      )}

      {/* AGENT-WISE UNIFIED COMPACT CARDS */}
      {loading ? (
        <div className="bg-white border border-slate-200 rounded-2xl py-10 text-center text-slate-500 text-xs sm:text-sm">
          <FaSpinner className="animate-spin inline mr-2" />
          Loading DOC tracking data...
        </div>
      ) : agentBlocksList.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-2xl py-10 text-center px-4">
          <p className="text-slate-600 text-xs sm:text-sm font-bold">
            {trackFilter !== "all"
              ? `No ${trackFilter} DOC for the selected filter`
              : fromDate || toDate || monthFilter
              ? "No DOC on this date"
              : "No DOC supply records found for the selected filter."}
          </p>
        </div>
      ) : (
        <div className="space-y-2.5">
          {agentBlocksList.map(({ agent, blocks, allBlocks }) => {
            const agentTotal = allBlocks.reduce((sum, bl) => sum + bl.total, 0);
            const agentTracked = allBlocks.reduce(
              (sum, bl) => sum + bl.tracked,
              0
            );
            const agentRemaining = Math.max(0, agentTotal - agentTracked);
            const isAgentAllTracked = agentTotal > 0 && agentRemaining <= 0;
            const showTotalBadge = allBlocks.length > 1;

            return (
              <div
                key={agent.id}
                className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden"
              >
                {/* AGENT HEADER WITH SINGLE "+" ICON ON MOBILE & TEXT ON DESKTOP */}
                <div className="bg-slate-100 border-b border-slate-200 px-2.5 sm:px-3.5 py-1.5 flex items-center justify-between gap-2">
                  <div className="flex items-center gap-1.5 min-w-0">
                    <span className="px-1.5 py-0.5 rounded bg-slate-200/90 border border-slate-300 text-slate-800 font-bold text-[10px] sm:text-xs shrink-0">
                      {agent.doc_id}
                    </span>
                    <h2 className="font-bold text-slate-800 text-xs sm:text-sm truncate">
                      {agent.agent_name}
                    </h2>
                    <span className="hidden sm:inline text-[11px] text-slate-500 shrink-0">
                      ({agent.upazila})
                    </span>
                  </div>

                  <div className="flex items-center gap-1.5 shrink-0">
                    {showTotalBadge && (
                      <span className="px-2 py-0.5 rounded bg-white border border-slate-300 text-slate-700 text-[10px] sm:text-xs font-semibold">
                        Total: <b>{agentTotal.toLocaleString()}</b>
                      </span>
                    )}

                    <button
                      type="button"
                      disabled={isAgentAllTracked}
                      title={
                        isAgentAllTracked
                          ? "All Tracked"
                          : `Add Tracking (${agentRemaining.toLocaleString()} left)`
                      }
                      onClick={() => openAgentTrackingModal(agent, allBlocks)}
                      className={`inline-flex items-center justify-center gap-1 w-7 h-7 sm:w-auto sm:h-auto sm:px-2.5 sm:py-1 rounded-full sm:rounded-lg text-[11px] sm:text-xs font-bold shadow-sm transition ${
                        isAgentAllTracked
                          ? "bg-slate-200 text-slate-500 cursor-not-allowed"
                          : "bg-emerald-600 hover:bg-emerald-700 text-white"
                      }`}
                    >
                      {isAgentAllTracked ? (
                        <>
                          <FaCheck size={11} />
                          <span className="hidden sm:inline">All Tracked</span>
                        </>
                      ) : (
                        <>
                          <FaPlus size={11} />
                          <span className="hidden sm:inline">
                            Add Tracking ({agentRemaining.toLocaleString()}{" "}
                            left)
                          </span>
                        </>
                      )}
                    </button>
                  </div>
                </div>

                {/* SINGLE UNIFIED SUPPLY + ULTRA-SLIM INLINE TRACKED TABLE */}
                <div className="overflow-x-auto">
                  <table className="w-full text-[11px] sm:text-xs border-collapse">
                    <thead className="bg-slate-50 border-b border-slate-200 text-slate-600">
                      <tr>
                        <th className="px-2 sm:px-3 py-1 text-left font-semibold whitespace-nowrap">
                          H.Date
                        </th>
                        <th className="px-2 sm:px-3 py-1 text-left font-semibold whitespace-nowrap">
                          DOC Type
                        </th>
                        <th className="px-2 sm:px-3 py-1 text-left font-semibold whitespace-nowrap">
                          AG/B+
                        </th>
                        <th className="px-2 sm:px-3 py-1 text-right font-semibold whitespace-nowrap">
                          Qty
                        </th>
                        <th className="px-2 sm:px-3 py-1 text-right font-semibold whitespace-nowrap">
                          B.Qty
                        </th>
                        <th className="px-2 sm:px-3 py-1 text-right font-semibold whitespace-nowrap">
                          Status
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {blocks.map((block) => {
                        const mapped = mapPcdToCategoryAndBreed(block.pcd);
                        const trackedRowsToShow =
                          trackFilter === "untracked" ? [] : block.trk;

                        return (
                          <React.Fragment key={`${block.h_date}_${block.pcd}`}>
                            {/* SUPPLY ROW */}
                            <tr className="bg-white hover:bg-slate-50/80">
                              <td className="px-2 sm:px-3 py-1 font-semibold text-slate-800 whitespace-nowrap">
                                {block.h_date}
                              </td>
                              <td className="px-2 sm:px-3 py-1 font-bold text-slate-800 whitespace-nowrap">
                                {block.pcd}{" "}
                                <span className="font-normal text-slate-500 text-[10px]">
                                  ({mapped.breed})
                                </span>
                              </td>
                              <td className="px-2 sm:px-3 py-1 text-slate-600 whitespace-nowrap">
                                {block.agBText}
                              </td>
                              <td className="px-2 sm:px-3 py-1 text-right font-extrabold text-slate-900 whitespace-nowrap">
                                {block.total.toLocaleString()}
                              </td>
                              <td className="px-2 sm:px-3 py-1 text-right text-slate-600 whitespace-nowrap">
                                {block.bTotal.toLocaleString()}
                              </td>
                              <td className="px-2 sm:px-3 py-1 text-right whitespace-nowrap">
                                {block.untracked <= 0 ? (
                                  <span className="px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold text-[10px]">
                                    All Tracked (
                                    {block.tracked.toLocaleString()})
                                  </span>
                                ) : (
                                  <span className="text-[10px] sm:text-[11px] font-semibold">
                                    <span className="text-emerald-700">
                                      Trk: {block.tracked.toLocaleString()}
                                    </span>{" "}
                                    •{" "}
                                    <span className="text-amber-700 font-bold">
                                      Left: {block.untracked.toLocaleString()}
                                    </span>
                                  </span>
                                )}
                              </td>
                            </tr>

                            {/* ULTRA-COMPACT SINGLE-LINE TRACKED FARMER ROWS (WITH TRACKED BY COLUMN) */}
                            {trackedRowsToShow.length > 0 && (
                              <tr className="bg-emerald-50/30">
                                <td colSpan={6} className="px-1.5 py-1">
                                  <div className="border border-emerald-200 rounded bg-white overflow-x-auto">
                                    <table className="w-full text-[10px] sm:text-[11px] leading-tight">
                                      <thead className="bg-emerald-50/80 border-b border-emerald-200 text-emerald-950">
                                        <tr className="whitespace-nowrap">
                                          <th className="py-0.5 px-1.5 text-left font-bold">
                                            Farmer
                                          </th>
                                          <th className="py-0.5 px-1.5 text-left font-bold">
                                            H.Date
                                          </th>
                                          <th className="py-0.5 px-1 text-left font-bold">
                                            AG/B+
                                          </th>
                                          <th className="py-0.5 px-1.5 text-right font-bold">
                                            DOC Qty
                                          </th>
                                          <th className="py-0.5 px-1.5 text-right font-bold">
                                            B.Qty
                                          </th>
                                          <th className="py-0.5 px-1.5 text-left font-bold">
                                            Feed
                                          </th>
                                          <th className="py-0.5 px-1.5 text-left font-bold">
                                            Breed
                                          </th>
                                          <th className="py-0.5 px-1.5 text-left font-bold">
                                            Mobile & Address
                                          </th>
                                          <th className="py-0.5 px-1.5 text-left font-bold">
                                            Tracked By
                                          </th>
                                          <th className="py-0.5 px-1.5 text-right font-bold">
                                            Act
                                          </th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-emerald-100">
                                        {trackedRowsToShow.map(
                                          (tr: DocTracking) => {
                                            const allowedToEdit =
                                              canModifyTracking(tr);
                                            const autoInfo =
                                              mapPcdToCategoryAndBreed(tr.pcd);
                                            const computedBox =
                                              tr.b_qty !== undefined &&
                                              tr.b_qty !== null
                                                ? Number(tr.b_qty)
                                                : block.docPerBox > 0
                                                ? Math.round(
                                                    Number(tr.qty || 0) /
                                                      block.docPerBox
                                                  )
                                                : "-";

                                            return (
                                              <tr
                                                key={tr.id}
                                                className="hover:bg-emerald-50/40 whitespace-nowrap"
                                              >
                                                <td className="py-0.5 px-1.5 font-bold text-slate-800">
                                                  {tr.farmer_name}
                                                </td>
                                                <td className="py-0.5 px-1.5 text-slate-600">
                                                  {tr.h_date}
                                                </td>
                                                <td className="py-0.5 px-1">
                                                  {agBadge(tr.ag_b) || "-"}
                                                </td>
                                                <td className="py-0.5 px-1.5 text-right font-extrabold text-emerald-700">
                                                  {Number(
                                                    tr.qty
                                                  ).toLocaleString()}
                                                </td>
                                                <td className="py-0.5 px-1.5 text-right font-semibold text-slate-700">
                                                  {computedBox}
                                                </td>
                                                <td className="py-0.5 px-1.5 font-semibold text-emerald-800">
                                                  {tr.feed_company || "-"}
                                                </td>
                                                <td className="py-0.5 px-1.5 text-slate-700">
                                                  {tr.category ||
                                                    autoInfo.category}{" "}
                                                  • {tr.breed || autoInfo.breed}
                                                </td>
                                                <td className="py-0.5 px-1.5 text-slate-600">
                                                  {[tr.mobile, tr.address]
                                                    .filter(Boolean)
                                                    .join(" • ") || "-"}
                                                </td>
                                                <td className="py-0.5 px-1.5 font-semibold text-indigo-700">
                                                  {tr.tracked_by_name || "-"}
                                                </td>
                                                <td className="py-0.5 px-1.5 text-right space-x-1">
                                                  {allowedToEdit ? (
                                                    <>
                                                      <button
                                                        type="button"
                                                        onClick={() =>
                                                          openEditTrackingModal(
                                                            tr,
                                                            agent
                                                          )
                                                        }
                                                        className="text-blue-600 hover:text-blue-800 p-0.5"
                                                        title="Edit"
                                                      >
                                                        <FaEdit size={10} />
                                                      </button>
                                                      <button
                                                        type="button"
                                                        onClick={() =>
                                                          handleDeleteTracking(
                                                            tr
                                                          )
                                                        }
                                                        className="text-red-500 hover:text-red-700 p-0.5"
                                                        title="Delete"
                                                      >
                                                        <FaTrash size={9} />
                                                      </button>
                                                    </>
                                                  ) : (
                                                    <span className="text-[9px] text-slate-400">
                                                      -
                                                    </span>
                                                  )}
                                                </td>
                                              </tr>
                                            );
                                          }
                                        )}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* MODAL 1: ADD / EDIT FARMER TRACKING */}
      {isTrackModalOpen && activeTarget && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-white rounded-t-2xl sm:rounded-2xl shadow-2xl w-full max-w-md max-h-[92vh] flex flex-col overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 border-b border-slate-200 bg-slate-100">
              <div>
                <h3 className="text-sm sm:text-base font-bold text-slate-800">
                  {editingTracking ? "Edit Tracking" : "Add Tracking"} —{" "}
                  {activeTarget.agent_name}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => {
                  setIsTrackModalOpen(false);
                  setEditingTracking(null);
                  setFarmerEntryMode("");
                }}
                className="w-7 h-7 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes />
              </button>
            </div>

            <form
              onSubmit={handleSaveTracking}
              className="p-4 space-y-3 overflow-y-auto text-xs"
            >
              {/* H.Date, DOC Type & AG/B+ Selectors */}
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    H.Date *
                  </label>
                  <select
                    value={trackForm.h_date || activeTarget.h_date}
                    disabled={!!editingTracking}
                    onChange={(e) => handleModalDateChange(e.target.value)}
                    className="w-full border border-slate-300 rounded-lg px-2 py-1.5 bg-white outline-none focus:border-emerald-500 disabled:bg-slate-100 font-semibold"
                  >
                    {(modalHatchDatesForAgent().length > 0
                      ? modalHatchDatesForAgent()
                      : [activeTarget.h_date]
                    ).map((d) => (
                      <option key={d} value={d}>
                        {d}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    DOC Type *
                  </label>
                  <select
                    value={trackForm.pcd || activeTarget.pcd}
                    disabled={!!editingTracking}
                    onChange={(e) => handleModalPcdChange(e.target.value)}
                    className="w-full border border-slate-300 rounded-lg px-2 py-1.5 bg-white outline-none focus:border-emerald-500 disabled:bg-slate-100 font-semibold"
                  >
                    {(modalDocTypesForDate(
                      trackForm.h_date || activeTarget.h_date
                    ).length > 0
                      ? modalDocTypesForDate(
                          trackForm.h_date || activeTarget.h_date
                        )
                      : [activeTarget.pcd]
                    ).map((p) => (
                      <option key={p} value={p}>
                        {p} ({mapPcdToCategoryAndBreed(p).breed})
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    AG / B+
                  </label>
                  <select
                    value={trackForm.ag_b}
                    onChange={(e) =>
                      setTrackForm({ ...trackForm, ag_b: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2 py-1.5 bg-white outline-none focus:border-emerald-500"
                  >
                    <option value="">-- Select --</option>
                    {modalSelectedStats.agBOptions.map((o) => (
                      <option key={o} value={o}>
                        {o}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Live Summary Bar for Selected Date & DOC Type (Includes Box Rate) */}
              <div className="grid grid-cols-4 gap-1.5 bg-slate-50 border border-slate-200 rounded-xl p-2 text-center">
                <div>
                  <span className="text-[10px] text-slate-500 block">
                    Total DOC
                  </span>
                  <strong className="text-slate-900 text-xs">
                    {modalSelectedStats.total.toLocaleString()}
                  </strong>
                </div>
                <div>
                  <span className="text-[10px] text-slate-500 block">
                    Total Box
                  </span>
                  <strong className="text-slate-800 text-xs">
                    {modalSelectedStats.bTotal.toLocaleString()}
                    {modalSelectedStats.docPerBox > 0
                      ? ` (${Math.round(modalSelectedStats.docPerBox)}/bx)`
                      : ""}
                  </strong>
                </div>
                <div>
                  <span className="text-[10px] text-emerald-700 block">
                    Tracked
                  </span>
                  <strong className="text-emerald-700 text-xs">
                    {modalSelectedStats.tracked.toLocaleString()}
                  </strong>
                </div>
                <div>
                  <span className="text-[10px] text-amber-700 block">
                    Remaining
                  </span>
                  <strong className="text-amber-700 text-xs">
                    {modalSelectedStats.remaining.toLocaleString()}
                  </strong>
                </div>
              </div>

              {/* 4 EXCLUSIVE SELECTION OPTIONS (When adding new tracking) */}
              {!editingTracking && (
                <div className="space-y-2 bg-slate-50/90 border border-slate-200 rounded-xl p-2.5">
                  {/* OPTION 1: Select farmer from this agent */}
                  <div>
                    <label className="block font-bold text-emerald-800 mb-1">
                      Select farmer from this agent
                    </label>
                    <select
                      value={
                        farmerEntryMode === "this_agent" ? selectedFarmerId : ""
                      }
                      onChange={(e) => {
                        const val = e.target.value;
                        if (!val) {
                          setFarmerEntryMode("");
                          setSelectedFarmerId("");
                          return;
                        }
                        setFarmerEntryMode("this_agent");
                        setSaveToFarmers(false);
                        setShowInVisits(true);
                        handleSelectFarmerFromDropdown(val);
                      }}
                      className="w-full border border-emerald-300 rounded-lg px-2.5 py-1.5 bg-white text-slate-800 outline-none focus:border-emerald-600 font-medium"
                    >
                      <option value="">
                        -- Select Farmer ({collaboratedFarmers.length} farmers) --
                      </option>
                      {collaboratedFarmers.map((f) => (
                        <option key={f.id} value={String(f.id)}>
                          {f.farmer_name}
                          {f.village ? ` - ${f.village}` : ""}
                          {f.mobile ? ` (${f.mobile})` : ""}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* CHECKBOX OPTIONS FOR OTHER 3 MODES */}
                  <div className="space-y-1.5 pt-1.5 border-t border-slate-200">
                    {/* OPTION 2: Select from others agent */}
                    <label className="flex items-center gap-2 font-semibold text-slate-700 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={farmerEntryMode === "other_agent"}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setFarmerEntryMode("other_agent");
                            setSelectedFarmerId("");
                            setSaveToFarmers(false);
                            setShowInVisits(true);
                          } else {
                            setFarmerEntryMode("");
                            setSelectedFarmerId("");
                          }
                        }}
                        className="w-3.5 h-3.5 accent-blue-600 rounded"
                      />
                      <span>Select from others agent</span>
                    </label>

                    {/* Compact Filters for "Select from others agent" */}
                    {farmerEntryMode === "other_agent" && (
                      <div className="grid grid-cols-2 gap-1.5 bg-blue-50/60 border border-blue-200 rounded-lg p-2 ml-4 text-[11px]">
                        <select
                          value={otherCompanyId}
                          onChange={(e) => {
                            setOtherCompanyId(e.target.value);
                            setOtherAgentId("");
                            setSelectedFarmerId("");
                          }}
                          className="border border-slate-300 rounded px-1.5 py-1 bg-white outline-none"
                        >
                          <option value="">All Companies</option>
                          {companiesList.map((c) => (
                            <option key={c.id} value={String(c.id)}>
                              {c.company_name}
                            </option>
                          ))}
                        </select>

                        <select
                          value={otherUpazila}
                          onChange={(e) => {
                            setOtherUpazila(e.target.value);
                            setOtherAgentId("");
                            setSelectedFarmerId("");
                          }}
                          className="border border-slate-300 rounded px-1.5 py-1 bg-white outline-none"
                        >
                          <option value="">All Upazilas</option>
                          {allFeedUpazilas.map((u) => (
                            <option key={u} value={u}>
                              {u}
                            </option>
                          ))}
                        </select>

                        <select
                          value={otherAgentId}
                          onChange={(e) => {
                            setOtherAgentId(e.target.value);
                            setSelectedFarmerId("");
                          }}
                          className="col-span-2 border border-slate-300 rounded px-1.5 py-1 bg-white outline-none font-semibold"
                        >
                          <option value="">
                            -- Select Other Agent ({otherAgentsOptions.length}) --
                          </option>
                          {otherAgentsOptions.map((a) => (
                            <option key={a.id} value={String(a.id)}>
                              {a.agent_name} ({a.upazila || "-"})
                            </option>
                          ))}
                        </select>

                        <select
                          value={selectedFarmerId}
                          disabled={!otherAgentId}
                          onChange={(e) =>
                            handleSelectFarmerFromDropdown(e.target.value)
                          }
                          className="col-span-2 border border-blue-300 rounded px-1.5 py-1 bg-white outline-none font-semibold text-blue-900 disabled:bg-slate-100"
                        >
                          <option value="">
                            -- Select Farmer ({otherAgentFarmersList.length}) --
                          </option>
                          {otherAgentFarmersList.map((f) => (
                            <option key={f.id} value={String(f.id)}>
                              {f.farmer_name}
                              {f.village ? ` - ${f.village}` : ""}
                              {f.mobile ? ` (${f.mobile})` : ""}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}

                    {/* OPTION 3: Create farmer & Track DOC */}
                    <label className="flex items-center gap-2 font-semibold text-indigo-800 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={farmerEntryMode === "create_farmer"}
                        onChange={(e) => {
                          if (e.target.checked) {
                            const mapped = mapPcdToCategoryAndBreed(
                              trackForm.pcd || activeTarget.pcd
                            );
                            setFarmerEntryMode("create_farmer");
                            setSelectedFarmerId("");
                            setSaveToFarmers(true);
                            setShowInVisits(true); // Default checked: Run this farmer
                            initNewFarmerFormDefaults();
                            setTrackForm((prev) => ({
                              ...prev,
                              farmer_name: "",
                              address: "",
                              mobile: "",
                              category: mapped.category,
                              breed: mapped.breed,
                            }));
                          } else {
                            setFarmerEntryMode("");
                            setSaveToFarmers(false);
                          }
                        }}
                        className="w-3.5 h-3.5 accent-indigo-600 rounded"
                      />
                      <span>Create farmer & Track DOC</span>
                    </label>

                    {/* OPTION 4: Only Track */}
                    <label className="flex items-center gap-2 font-semibold text-amber-800 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={farmerEntryMode === "only_track"}
                        onChange={(e) => {
                          if (e.target.checked) {
                            const mapped = mapPcdToCategoryAndBreed(
                              trackForm.pcd || activeTarget.pcd
                            );
                            setFarmerEntryMode("only_track");
                            setSelectedFarmerId("");
                            setSaveToFarmers(false);
                            setShowInVisits(false); // No running option for Only Track
                            setTrackForm((prev) => ({
                              ...prev,
                              farmer_name: "",
                              address: "",
                              mobile: "",
                              category: mapped.category,
                              breed: mapped.breed,
                            }));
                          } else {
                            setFarmerEntryMode("");
                          }
                        }}
                        className="w-3.5 h-3.5 accent-amber-600 rounded"
                      />
                      <span>Only Track</span>
                    </label>
                  </div>
                </div>
              )}

              {/* FULL FARMER ADD FORM WITH LOCATION PICKER (Shown when "Create farmer & Track DOC" is checked) */}
              {!editingTracking && farmerEntryMode === "create_farmer" && (
                <div className="border border-indigo-200 bg-indigo-50/40 rounded-xl p-2.5 space-y-2">
                  <div className="font-bold text-[11px] text-indigo-900 border-b border-indigo-200 pb-1">
                    Farmer Database Information
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block font-medium text-slate-700 mb-0.5">
                        Company *
                      </label>
                      <select
                        value={newFarmerForm.company_id}
                        onChange={(e) =>
                          setNewFarmerForm({
                            ...newFarmerForm,
                            company_id: e.target.value,
                            agent_id: "",
                          })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2 py-1 bg-white"
                      >
                        <option value="">Select</option>
                        {companiesList.map((c) => (
                          <option key={c.id} value={c.id}>
                            {c.company_name}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block font-medium text-slate-700 mb-0.5">
                        District *
                      </label>
                      <select
                        value={newFarmerForm.district}
                        onChange={(e) =>
                          setNewFarmerForm({
                            ...newFarmerForm,
                            district: e.target.value,
                            upazila: "",
                            union_name: "",
                            agent_id: "",
                          })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2 py-1 bg-white"
                      >
                        <option value="">Select</option>
                        {geoDistricts.map((d) => (
                          <option key={d.id} value={d.name}>
                            {d.name}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block font-medium text-slate-700 mb-0.5">
                        Upazila *
                      </label>
                      <select
                        value={newFarmerForm.upazila}
                        onChange={(e) =>
                          setNewFarmerForm({
                            ...newFarmerForm,
                            upazila: e.target.value,
                            union_name: "",
                            agent_id: "",
                          })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2 py-1 bg-white"
                      >
                        <option value="">Select</option>
                        {geoUpazilas.map((u) => (
                          <option key={u.id} value={u.name}>
                            {u.name}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block font-medium text-slate-700 mb-0.5">
                        Agent *
                      </label>
                      <select
                        value={newFarmerForm.agent_id}
                        onChange={(e) =>
                          setNewFarmerForm({
                            ...newFarmerForm,
                            agent_id: e.target.value,
                          })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2 py-1 bg-white"
                      >
                        <option value="">Select</option>
                        {nfAgentOptions.map((a) => (
                          <option key={a.id} value={a.id}>
                            {a.agent_name}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block font-medium text-slate-700 mb-0.5">
                        Union
                      </label>
                      <select
                        value={newFarmerForm.union_name}
                        onChange={(e) =>
                          setNewFarmerForm({
                            ...newFarmerForm,
                            union_name: e.target.value,
                          })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2 py-1 bg-white"
                      >
                        <option value="">Select</option>
                        {geoUnions.map((u) => (
                          <option key={u.id} value={u.name}>
                            {u.name}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block font-medium text-slate-700 mb-0.5">
                        Village
                      </label>
                      <input
                        placeholder="Village"
                        value={newFarmerForm.village}
                        onChange={(e) => {
                          const v = e.target.value;
                          setNewFarmerForm({
                            ...newFarmerForm,
                            village: v,
                          });
                          setTrackForm((prev) => ({
                            ...prev,
                            address: v,
                          }));
                        }}
                        className="w-full border border-slate-300 rounded-lg px-2 py-1 bg-white"
                      />
                    </div>
                  </div>

                  {/* Map Location Picker Trigger */}
                  <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-indigo-200 rounded-lg p-2">
                    <button
                      type="button"
                      onClick={() => {
                        setPickerCenterOverride(null);
                        setShowCreateFarmerMap(true);
                      }}
                      className="inline-flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold"
                    >
                      <MapPin size={12} /> Select Location from Map
                    </button>
                    {newFarmerForm.latitude && newFarmerForm.longitude && (
                      <span className="text-[11px] font-semibold text-emerald-700">
                        📍 ({Number(newFarmerForm.latitude).toFixed(4)},{" "}
                        {Number(newFarmerForm.longitude).toFixed(4)})
                      </span>
                    )}
                  </div>
                </div>
              )}

              {/* TRACKING FORM FIELDS (Shown ONLY after an option/farmer is selected!) */}
              {isTrackingFormVisible && (
                <div className="space-y-3 pt-2 border-t border-slate-200">
                  {/* Farmer Name with Floating Duplicate Suggestion Dropdown */}
                  <div className="relative">
                    <label className="block font-medium text-slate-700 mb-1">
                      Farmer Name *
                    </label>
                    <input
                      type="text"
                      required
                      placeholder="Enter farmer name"
                      value={trackForm.farmer_name}
                      onFocus={() => setShowSimilarSuggestions(true)}
                      onBlur={() =>
                        setTimeout(() => setShowSimilarSuggestions(false), 220)
                      }
                      onChange={(e) => {
                        setTrackForm({
                          ...trackForm,
                          farmer_name: e.target.value,
                        });
                        setShowSimilarSuggestions(true);
                      }}
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                    />

                    {/* Floating Similarity Suggestions */}
                    {showSimilarSuggestions &&
                      similarFarmersList.length > 0 && (
                        <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-amber-300 rounded-xl shadow-xl z-50 max-h-44 overflow-y-auto divide-y divide-slate-100">
                          <div className="bg-amber-50 px-2.5 py-1 text-[10px] font-bold text-amber-900 flex items-center justify-between">
                            <span>
                              ⚠️ Similar Farmers Found (Click to select)
                            </span>
                            <button
                              type="button"
                              onClick={() => setShowSimilarSuggestions(false)}
                              className="text-slate-400 hover:text-slate-700"
                            >
                              ✕
                            </button>
                          </div>
                          {similarFarmersList.map(
                            ({ farmer: sf, agentName, companyName }) => (
                              <button
                                key={sf.id}
                                type="button"
                                onMouseDown={() =>
                                  handlePickSuggestedFarmer(sf)
                                }
                                className="w-full text-left px-2.5 py-1.5 hover:bg-emerald-50 transition block"
                              >
                                <div className="font-bold text-slate-800 text-xs flex items-center justify-between">
                                  <span>{sf.farmer_name}</span>
                                  {sf.mobile && (
                                    <span className="text-[10px] font-normal text-slate-500">
                                      📞 {sf.mobile}
                                    </span>
                                  )}
                                </div>
                                <div className="text-[10px] text-slate-500 truncate">
                                  Agent: <b>{agentName}</b> • {companyName} •{" "}
                                  {[sf.village, sf.upazila]
                                    .filter(Boolean)
                                    .join(", ") || "-"}
                                </div>
                              </button>
                            )
                          )}
                        </div>
                      )}
                  </div>

                  <div className="grid grid-cols-2 gap-2.5">
                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        Address
                      </label>
                      <input
                        type="text"
                        placeholder="Village / Area"
                        value={trackForm.address}
                        onChange={(e) => {
                          const val = e.target.value;
                          setTrackForm({ ...trackForm, address: val });
                          if (farmerEntryMode === "create_farmer") {
                            setNewFarmerForm((prev) => ({
                              ...prev,
                              village: val,
                            }));
                          }
                        }}
                        className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                      />
                    </div>

                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        Mobile
                      </label>
                      <input
                        type="text"
                        placeholder="017XXXXXXXX"
                        value={trackForm.mobile}
                        onChange={(e) =>
                          setTrackForm({ ...trackForm, mobile: e.target.value })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-2.5">
                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        Farm Category *
                      </label>
                      <select
                        value={trackForm.category}
                        onChange={(e) => {
                          const cat = e.target.value;
                          const breeds = BREEDS_BY_CATEGORY[cat] || [];
                          setTrackForm({
                            ...trackForm,
                            category: cat,
                            breed: breeds[0] || "",
                          });
                        }}
                        className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white outline-none focus:border-emerald-500"
                      >
                        {FARM_CATEGORIES.map((cat) => (
                          <option key={cat} value={cat}>
                            {cat}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        Breed ({trackForm.pcd || activeTarget.pcd}) *
                      </label>
                      <select
                        value={trackForm.breed}
                        onChange={(e) =>
                          setTrackForm({ ...trackForm, breed: e.target.value })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white outline-none focus:border-emerald-500"
                      >
                        {(BREEDS_BY_CATEGORY[trackForm.category] || []).map(
                          (b) => (
                            <option key={b} value={b}>
                              {b}
                            </option>
                          )
                        )}
                      </select>
                    </div>
                  </div>

                  {/* DOC Qty, Auto Box Qty (B.Qty), and Feed Company */}
                  <div className="grid grid-cols-3 gap-2">
                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        DOC Qty *
                      </label>
                      <input
                        type="number"
                        required
                        min={1}
                        max={modalSelectedStats.remaining}
                        placeholder="Qty"
                        value={trackForm.qty}
                        onChange={(e) => {
                          const val = e.target.value;
                          const autoBx = computeAutoBoxQty(
                            val,
                            modalSelectedStats.docPerBox
                          );
                          setTrackForm({
                            ...trackForm,
                            qty: val,
                            box_qty: autoBx,
                          });
                        }}
                        className="w-full border border-slate-300 rounded-lg px-2 py-1.5 outline-none focus:border-emerald-500 font-bold"
                      />
                    </div>

                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        Box Qty
                      </label>
                      <input
                        type="number"
                        min={0}
                        placeholder="Box"
                        value={trackForm.box_qty}
                        onChange={(e) =>
                          setTrackForm({
                            ...trackForm,
                            box_qty: e.target.value,
                          })
                        }
                        className="w-full border border-emerald-300 bg-emerald-50/40 rounded-lg px-2 py-1.5 outline-none focus:border-emerald-600 font-bold text-emerald-800"
                      />
                    </div>

                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        Feed Company *
                      </label>
                      <select
                        required
                        value={trackForm.feed_company}
                        onChange={(e) =>
                          setTrackForm({
                            ...trackForm,
                            feed_company: e.target.value,
                          })
                        }
                        className="w-full border border-slate-300 rounded-lg px-2 py-1.5 outline-none bg-white focus:border-emerald-500"
                      >
                        <option value="">Select</option>
                        {companyNames.map((fc) => (
                          <option key={fc} value={fc}>
                            {fc}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {Number(trackForm.qty) > 0 &&
                    Number(trackForm.box_qty) > 0 && (
                      <div className="text-[11px] text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-lg px-2.5 py-1">
                        Tracking DOC: <b>{Number(trackForm.qty)}</b> •
                        Running/Farmer Stock (+{Number(trackForm.box_qty)} bonus
                        for {Number(trackForm.box_qty)} box):{" "}
                        <b>
                          {Number(trackForm.qty) + Number(trackForm.box_qty)}{" "}
                          Pcs
                        </b>
                      </div>
                    )}

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">
                      Remark
                    </label>
                    <input
                      type="text"
                      placeholder="Optional note..."
                      value={trackForm.remark}
                      onChange={(e) =>
                        setTrackForm({ ...trackForm, remark: e.target.value })
                      }
                      className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Batch Add or Replace Option (Hidden in Only Track mode) */}
                  {!editingTracking &&
                    farmerEntryMode !== "only_track" &&
                    batchCtx &&
                    needsBatchChoice && (
                      <div className="border border-amber-300 bg-amber-50 rounded-xl p-2.5 space-y-1.5">
                        <div className="font-bold text-amber-900">
                          Before: {batchCtx.prevCats.join(", ")}{" "}
                          {batchCtx.inactive ? "(inactive)" : "(running)"} →
                          Now: {trackForm.category}
                        </div>
                        <label className="flex items-start gap-2 cursor-pointer">
                          <input
                            type="radio"
                            name="batchMode"
                            checked={batchMode === "add"}
                            onChange={() => setBatchMode("add")}
                            className="mt-0.5 accent-amber-600"
                          />
                          <span>
                            <b>Add</b> new {trackForm.category} batch (keep{" "}
                            {batchCtx.prevCats.join(", ")})
                          </span>
                        </label>
                        <label className="flex items-start gap-2 cursor-pointer">
                          <input
                            type="radio"
                            name="batchMode"
                            checked={batchMode === "replace"}
                            onChange={() => setBatchMode("replace")}
                            className="mt-0.5 accent-amber-600"
                          />
                          <span>
                            <b>Replace</b> {batchCtx.prevCats.join(", ")} with{" "}
                            {trackForm.category}
                          </span>
                        </label>
                      </div>
                    )}

                  {/* Run this farmer / Show as Running on Visit Page (Hidden in "Only Track" mode) */}
                  {!editingTracking && farmerEntryMode !== "only_track" && (
                    <div className="space-y-2 pt-1">
                      <label className="flex items-center gap-2 font-semibold text-emerald-800 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={showInVisits}
                          onChange={(e) => setShowInVisits(e.target.checked)}
                          className="w-4 h-4 accent-emerald-600"
                        />
                        {farmerEntryMode === "create_farmer"
                          ? "Run this farmer (Show as Running on Visit page)"
                          : "Show as Running on Visit page"}
                      </label>

                      {showInVisits && (
                        <div>
                          <label className="block font-medium text-slate-700 mb-1">
                            Batch mark (optional)
                          </label>
                          <input
                            type="text"
                            placeholder="e.g. Shed-2 / Sonali Batch (1)"
                            value={trackForm.batch_mark}
                            onChange={(e) =>
                              setTrackForm({
                                ...trackForm,
                                batch_mark: e.target.value,
                              })
                            }
                            className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                          />
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setIsTrackModalOpen(false);
                    setEditingTracking(null);
                    setFarmerEntryMode("");
                  }}
                  className="flex-1 py-2 rounded-xl border border-slate-300 text-slate-600 font-semibold"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={
                    submitting ||
                    modalSelectedStats.remaining <= 0 ||
                    !isTrackingFormVisible
                  }
                  className="flex-1 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-bold flex items-center justify-center gap-1.5"
                >
                  {submitting && <FaSpinner className="animate-spin" />}
                  {editingTracking ? "Update Tracking" : "Save Tracking"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* LOCATION PICKER MODAL FOR "CREATE FARMER & TRACK DOC" */}
      {showCreateFarmerMap && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[2100] p-2 sm:p-4">
          <div className="bg-white p-3 sm:p-4 rounded-xl w-full max-w-4xl shadow-2xl space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="font-bold text-sm sm:text-base">
                Select Farm Location
              </h2>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleLocateMeInNewFarmerPicker}
                  disabled={locatingInPicker}
                  className="inline-flex items-center gap-1 bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                >
                  <Navigation size={13} />
                  {locatingInPicker ? "Locating..." : "Locate Me"}
                </button>

                {newFarmerForm.latitude && newFarmerForm.longitude && (
                  <button
                    type="button"
                    onClick={() => setShowCreateFarmerMap(false)}
                    className="inline-flex items-center gap-1 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                  >
                    <CheckCircle2 size={13} /> Use Selected Location
                  </button>
                )}

                <button
                  type="button"
                  onClick={() => setShowCreateFarmerMap(false)}
                >
                  <X size={18} />
                </button>
              </div>
            </div>

            <LocationPicker
              key={
                pickerCenterOverride
                  ? `${pickerCenterOverride[0]}_${pickerCenterOverride[1]}`
                  : "doc_new_farmer_picker"
              }
              center={
                pickerCenterOverride
                  ? pickerCenterOverride
                  : selectedNewFarmerUpazilaObj?.lat &&
                    selectedNewFarmerUpazilaObj?.lon
                  ? [
                      Number(selectedNewFarmerUpazilaObj.lat),
                      Number(selectedNewFarmerUpazilaObj.lon),
                    ]
                  : selectedNewFarmerDistrictObj?.lat &&
                    selectedNewFarmerDistrictObj?.lon
                  ? [
                      Number(selectedNewFarmerDistrictObj.lat),
                      Number(selectedNewFarmerDistrictObj.lon),
                    ]
                  : [24.7471, 90.4203]
              }
              zoom={
                pickerCenterOverride
                  ? 17
                  : selectedNewFarmerUpazilaObj
                  ? 13
                  : selectedNewFarmerDistrictObj
                  ? 11
                  : 8
              }
              selectedPosition={
                newFarmerForm.latitude && newFarmerForm.longitude
                  ? [
                      Number(newFarmerForm.latitude),
                      Number(newFarmerForm.longitude),
                    ]
                  : null
              }
              onSelect={(lat, lng) => {
                setNewFarmerForm((prev) => ({
                  ...prev,
                  latitude: lat.toString(),
                  longitude: lng.toString(),
                }));
                setShowCreateFarmerMap(false);
              }}
            />
          </div>
        </div>
      )}

      {/* MODAL 2: EXPORT EXCEL */}
      {isExportModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-sm overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 border-b border-slate-200 bg-slate-100">
              <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <FaFileExcel className="text-emerald-600" />
                Select DOC Types to Export
              </h3>
              <button
                onClick={() => setIsExportModalOpen(false)}
                className="w-7 h-7 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes size={13} />
              </button>
            </div>

            <div className="p-4 space-y-3">
              <div className="flex items-center justify-between text-xs">
                <span className="text-slate-500">
                  Date: <strong>{dateRangeLabel}</strong>
                </span>
                <button
                  type="button"
                  onClick={() =>
                    setExportSelectedTypes(
                      exportSelectedTypes.length ===
                        categoryFilteredPcdTypes.length
                        ? []
                        : [...categoryFilteredPcdTypes]
                    )
                  }
                  className="text-emerald-600 font-semibold hover:underline"
                >
                  {exportSelectedTypes.length ===
                  categoryFilteredPcdTypes.length
                    ? "Deselect All"
                    : "Select All"}
                </button>
              </div>

              <div className="border border-slate-200 rounded-xl max-h-48 overflow-y-auto p-2 space-y-1 bg-slate-50">
                {categoryFilteredPcdTypes.map((type) => {
                  const checked = exportSelectedTypes.includes(type);
                  return (
                    <label
                      key={type}
                      className={`flex items-center gap-2.5 px-3 py-2 rounded-lg cursor-pointer text-xs font-medium transition ${
                        checked
                          ? "bg-emerald-50 text-emerald-800 border border-emerald-200"
                          : "hover:bg-white text-slate-700"
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleExportDocType(type)}
                        className="w-4 h-4 rounded text-emerald-600"
                      />
                      <span>
                        {type} ({mapPcdToCategoryAndBreed(type).breed})
                      </span>
                    </label>
                  );
                })}
              </div>

              <div className="flex gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setIsExportModalOpen(false)}
                  className="flex-1 py-2 rounded-xl border border-slate-300 text-slate-600 text-xs font-semibold"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleExportExcel}
                  disabled={exportSelectedTypes.length === 0}
                  className="flex-1 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-semibold flex items-center justify-center gap-1.5"
                >
                  <FaFileExcel />
                  Export ({exportSelectedTypes.length})
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: TRACKED REPORT POPUP */}
      {isReportOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-white rounded-t-2xl sm:rounded-2xl shadow-2xl w-full max-w-2xl max-h-[92vh] flex flex-col overflow-hidden">
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-slate-200 bg-slate-100">
              <div className="flex items-center gap-2">
                <FaChartPie className="text-amber-500" />
                <h3 className="text-base font-bold text-slate-800">
                  DOC Tracked Report
                </h3>
              </div>
              <button
                onClick={() => setIsReportOpen(false)}
                className="w-8 h-8 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes />
              </button>
            </div>

            <div className="p-4 sm:p-5 space-y-4 overflow-y-auto">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 bg-slate-50 p-3 rounded-xl border border-slate-200">
                <div>
                  <label className="block text-[11px] font-semibold text-slate-500 mb-1">
                    From Date
                  </label>
                  <input
                    type="date"
                    value={reportFromDate}
                    onChange={(e) => setReportFromDate(e.target.value)}
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs sm:text-sm bg-white"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-semibold text-slate-500 mb-1">
                    To Date
                  </label>
                  <input
                    type="date"
                    value={reportToDate}
                    onChange={(e) => setReportToDate(e.target.value)}
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs sm:text-sm bg-white"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-semibold text-slate-500 mb-1">
                    DOC Type
                  </label>
                  <select
                    value={reportPcdFilter}
                    onChange={(e) => setReportPcdFilter(e.target.value)}
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs sm:text-sm bg-white"
                  >
                    <option value="">All DOC Types</option>
                    {allPcdTypes.map((p) => (
                      <option key={p} value={p}>
                        {p} ({mapPcdToCategoryAndBreed(p).breed})
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-2.5">
                <div className="bg-slate-100 border border-slate-200 rounded-xl p-3 text-center">
                  <p className="text-xs text-slate-500">Total DOC</p>
                  <p className="text-base sm:text-lg font-bold text-slate-800">
                    {repTotalDoc.toLocaleString()}
                  </p>
                </div>

                <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 text-center">
                  <p className="text-xs text-emerald-700">Total Tracked</p>
                  <p className="text-base sm:text-lg font-bold text-emerald-700">
                    {repTotalTracked.toLocaleString()} ({repTrackedPct}%)
                  </p>
                </div>

                <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-center">
                  <p className="text-xs text-amber-700">Untracked</p>
                  <p className="text-base sm:text-lg font-bold text-amber-700">
                    {repTotalUntracked.toLocaleString()} ({repUntrackedPct}%)
                  </p>
                </div>
              </div>

              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">
                  Feed Company Wise Tracked Report
                </h4>
                <div className="border border-slate-200 rounded-xl overflow-hidden">
                  <table className="min-w-full text-xs sm:text-sm">
                    <thead className="bg-slate-50 border-b border-slate-200 text-slate-600">
                      <tr>
                        <th className="px-3.5 py-2 text-left font-semibold">
                          Feed Company
                        </th>
                        <th className="px-3.5 py-2 text-right font-semibold">
                          Tracked Qty
                        </th>
                        <th className="px-3.5 py-2 text-right font-semibold">
                          % of Tracked
                        </th>
                        <th className="px-3.5 py-2 text-right font-semibold">
                          % of Total DOC
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {companyBreakdown.map((item) => (
                        <tr key={item.company} className="hover:bg-slate-50">
                          <td className="px-3.5 py-2 font-semibold text-slate-800">
                            {item.company}
                          </td>
                          <td className="px-3.5 py-2 text-right font-bold text-emerald-700">
                            {item.qty.toLocaleString()}
                          </td>
                          <td className="px-3.5 py-2 text-right text-slate-600">
                            {item.pctOfTracked}%
                          </td>
                          <td className="px-3.5 py-2 text-right text-slate-500">
                            {item.pctOfTotal}%
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">
                  DOC Type Wise Summary
                </h4>
                <div className="border border-slate-200 rounded-xl overflow-hidden">
                  <table className="min-w-full text-xs sm:text-sm">
                    <thead className="bg-slate-50 border-b border-slate-200 text-slate-600">
                      <tr>
                        <th className="px-3.5 py-2 text-left font-semibold">
                          DOC Type
                        </th>
                        <th className="px-3.5 py-2 text-right font-semibold">
                          Total DOC
                        </th>
                        <th className="px-3.5 py-2 text-right font-semibold">
                          Tracked
                        </th>
                        <th className="px-3.5 py-2 text-right font-semibold">
                          Untracked
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {pcdBreakdown.map((p) => (
                        <tr key={p.pcd} className="hover:bg-slate-50">
                          <td className="px-3.5 py-2 font-bold text-slate-800">
                            {p.pcd} ({mapPcdToCategoryAndBreed(p.pcd).breed})
                          </td>
                          <td className="px-3.5 py-2 text-right font-semibold">
                            {p.supQty.toLocaleString()}
                          </td>
                          <td className="px-3.5 py-2 text-right text-emerald-700 font-semibold">
                            {p.trkQty.toLocaleString()} ({p.pct}%)
                          </td>
                          <td className="px-3.5 py-2 text-right text-amber-700 font-semibold">
                            {p.untrkQty.toLocaleString()}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex justify-end">
              <button
                onClick={() => setIsReportOpen(false)}
                className="px-5 py-2 rounded-xl bg-slate-800 text-white text-xs sm:text-sm font-medium"
              >
                Close Report
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}