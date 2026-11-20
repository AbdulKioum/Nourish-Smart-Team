import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Filter, ChevronDown, Stethoscope } from "lucide-react";
import { supabase, rawSupabase } from "../../lib/supabase";
import ExcelJS from "exceljs";
import SearchableSelect from "../../components/common/SearchableSelect";

import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Capacitor } from "@capacitor/core";

function calcAgeDays(hatchDate?: string, visitDate?: string): number {
  if (!hatchDate) return 0;
  const hd = new Date(hatchDate);
  const vd = visitDate ? new Date(visitDate) : new Date();
  return Math.max(0, Math.floor((vd.getTime() - hd.getTime()) / 86400000));
}

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

function getRemarkText(rep: any): string {
  const r = rep?.remarks ?? rep?.remark ?? "";
  return String(r || "").trim();
}

function formatDmy(dateStr?: string) {
  if (!dateStr) return "-";
  const parts = String(dateStr).slice(0, 10).split("-");
  if (parts.length !== 3) return dateStr;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

const STOCK_FIELDS = [
  "broiler","layer","sonali","duck","quail","deshi_chicken",
  "dairy_cow","bull","calf_heifer",
  "pangus","tilapia","koi","shing","pabda","gulsha","magur","carp","golda","bagda","tengra",
];

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

let cachedPageState: {
  initialized: boolean;
  userRole: string;
  assignedUpazilas: string[];
  assignedDistricts: string[];
  agents: any[];
  farmers: any[];
  subAgents: any[];
  visitsMap: Record<string, any>;
  beforeReports: any[];
  afterReports: any[];
  layerReports: any[];
  broilerStds: Record<number, any>;
  sonaliStds: Record<number, any>;
  layerStds: Record<string, any>;
  companies: any[];
  usersList: any[];
  currentUserId: string;
  userProfile: { fullName: string; designation: string; region: string };
  companyFilter: string;
  selectedCategory: "all" | "Broiler" | "Layer" | "Sonali" | "Fish" | "Cattle";
  reportType: "all" | "before_sale" | "after_sale";
  userFilter: string;
  upazilaFilter: string;
  agentFilter: string;
  farmerFilter: string;
  feedTypeFilter: string;
  isMultiAge: boolean;
  ageFilter: string;
  ageFrom: string;
  ageTo: string;
  monthFilter: string;
  isMultiDate: boolean;
  fromDate: string;
  toDate: string;
} | null = null;

export default function ReportsPage() {
  const navigate = useNavigate();
  const [selectedCategory, setSelectedCategory] = useState<
    "all" | "Broiler" | "Layer" | "Sonali" | "Fish" | "Cattle"
  >(cachedPageState?.selectedCategory || "all");
  const [reportType, setReportType] = useState<
    "all" | "before_sale" | "after_sale"
  >(cachedPageState?.reportType || "all");

  const [userRole, setUserRole] = useState<string>(
    cachedPageState?.userRole || "officer"
  );
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>(
    cachedPageState?.assignedUpazilas || []
  );
  const [_assignedDistricts, setAssignedDistricts] = useState<string[]>(
    cachedPageState?.assignedDistricts || []
  );
  const [currentUserId, setCurrentUserId] = useState<string>(
    cachedPageState?.currentUserId || ""
  );
  // Footer lines for the WhatsApp message (same defaults as VisitsPage).
  const [currentUserProfile, setCurrentUserProfile] = useState(
    cachedPageState?.userProfile || {
      fullName: "Officer",
      designation: "Customer Service Officer",
      region: "Mymensingh Zone",
    }
  );
  const [usersList, setUsersList] = useState<any[]>(
    cachedPageState?.usersList || []
  );

  const [agents, setAgents] = useState<any[]>(cachedPageState?.agents || []);
  const [farmers, setFarmers] = useState<any[]>(
    cachedPageState?.farmers || []
  );
  const [subAgents, setSubAgents] = useState<any[]>(
    cachedPageState?.subAgents || []
  );
  const [companies, setCompanies] = useState<any[]>(
    cachedPageState?.companies || []
  );
  const [visitsMap, setVisitsMap] = useState<Record<string, any>>(
    cachedPageState?.visitsMap || {}
  );
  const [beforeReports, setBeforeReports] = useState<any[]>(
    cachedPageState?.beforeReports || []
  );
  const [afterReports, setAfterReports] = useState<any[]>(
    cachedPageState?.afterReports || []
  );
  const [layerReports, setLayerReports] = useState<any[]>(
    cachedPageState?.layerReports || []
  );
  const [broilerStds, setBroilerStds] = useState<Record<number, any>>(
    cachedPageState?.broilerStds || {}
  );
  const [sonaliStds, setSonaliStds] = useState<Record<number, any>>(
    cachedPageState?.sonaliStds || {}
  );
  const [layerStds, setLayerStds] = useState<Record<string, any>>(
    cachedPageState?.layerStds || {}
  );
  const [loading, setLoading] = useState(!cachedPageState?.initialized);

  // Filters
  const [companyFilter, setCompanyFilter] = useState<string>(
    cachedPageState?.companyFilter || "Nourish"
  );
  const [upazilaFilter, setUpazilaFilter] = useState<string>(
    cachedPageState?.upazilaFilter || "all"
  );
  const [agentFilter, setAgentFilter] = useState<string>(
    cachedPageState?.agentFilter || "all"
  );
  const [farmerFilter, setFarmerFilter] = useState<string>(
    cachedPageState?.farmerFilter || "all"
  );
  const [feedTypeFilter, setFeedTypeFilter] = useState<string>(
    cachedPageState?.feedTypeFilter || "all"
  );
  const [isMultiAge, setIsMultiAge] = useState<boolean>(
    cachedPageState?.isMultiAge || false
  );
  const [ageFilter, setAgeFilter] = useState<string>(
    cachedPageState?.ageFilter || "all"
  );
  const [ageFrom, setAgeFrom] = useState<string>(
    cachedPageState?.ageFrom || ""
  );
  const [ageTo, setAgeTo] = useState<string>(cachedPageState?.ageTo || "");
  const [monthFilter, setMonthFilter] = useState<string>(
    cachedPageState?.monthFilter || ""
  );
  const [isMultiDate, setIsMultiDate] = useState<boolean>(
    cachedPageState?.isMultiDate || false
  );
  const [fromDate, setFromDate] = useState<string>(
    cachedPageState?.fromDate || ""
  );
  const [toDate, setToDate] = useState<string>(
    cachedPageState?.toDate || ""
  );
  const [userFilter, setUserFilter] = useState<string>(
    cachedPageState?.userFilter || "all"
  );

  // WhatsApp Popup State (same as VisitsPage)
  const [whatsappPopupText, setWhatsappPopupText] = useState<string>("");
  const [copiedWa, setCopiedWa] = useState(false);

  // Mobile filter collapse (UI-only, not cached): on mobile the Top Bar
  // and Filter Bar hide behind one Filter button; on PC (md+) they stay
  // open exactly as before.
  const [showFilters, setShowFilters] = useState(false);

  // Selected report row (UI-only, not cached): tapping a row highlights it
  // so it stays identifiable while scrolling left-right on the phone.
  const [selectedReportKey, setSelectedReportKey] = useState<string | null>(
    null
  );

  const isFullAccess = userRole === "super_admin" || userRole === "manager";

  useEffect(() => {
    (async () => {
      if (cachedPageState?.initialized) {
        // The module cache survives a logout/login in the same tab — never
        // serve another user's cached page (role, upazilas, users list).
        // If the session user changed, drop the cache and load fresh.
        // An empty cached users list is also never trusted: the list
        // always has at least the current user, so an empty one means a
        // stale failed load — drop the cache so the network retry runs.
        const cachedUserId = cachedPageState.currentUserId;
        const cachedUsersEmpty =
          (cachedPageState.usersList || []).length === 0;
        const {
          data: { user },
        } = await supabase.auth.getUser();
        if (
          (user && String(cachedUserId || "") !== String(user.id)) ||
          cachedUsersEmpty
        ) {
          cachedPageState = null;
          loadReportsData();
        }
        return;
      }
      loadReportsData();
    })();
  }, []);

  async function loadReportsData() {
    setLoading(true);

    const {
      data: { user },
    } = await supabase.auth.getUser();

    let role = "officer";
    let uUpazilas: string[] = [];
    let uDistricts: string[] = [];

    // Hoisted so the cache save at the end of this load can use it.
    let uProfile = {
      fullName: "Officer",
      designation: "Customer Service Officer",
      region: "Mymensingh Zone",
    };
    let uid = "";

    if (user) {
      uid = user.id;
      setCurrentUserId(user.id);
      // Full name / designation / region feed the WhatsApp message
      // footer (same as VisitsPage). If those columns are missing, fall
      // back to the plain select so the load never breaks.
      let uProf: any = null;
      const { data: uProfData, error: uProfErr } = await supabase
        .from("users")
        .select(
          "role, upazilas, districts, full_name, email, designation, region"
        )
        .eq("id", user.id)
        .single();
      if (!uProfErr && uProfData) {
        uProf = uProfData;
      } else {
        const retry = await supabase
          .from("users")
          .select("role, upazilas, districts")
          .eq("id", user.id)
          .single();
        uProf = retry.data || null;
      }

      if (uProf) {
        role = uProf.role || "officer";
        uUpazilas = parseList(uProf.upazilas);
        uDistricts = parseList(uProf.districts);
        uProfile = {
          fullName:
            uProf.full_name || uProf.email?.split("@")[0] || "Officer",
          designation:
            uProf.designation || "Customer Service Officer (Poultry)",
          region: uProf.region || "Mymensingh Zone",
        };
        setCurrentUserProfile(uProfile);
      }
    }

    setUserRole(role);
    setAssignedUpazilas(uUpazilas);
    setAssignedDistricts(uDistricts);

    const fullAccess = role === "super_admin" || role === "manager";
    if (!fullAccess && uUpazilas.length === 1) {
      setUpazilaFilter(uUpazilas[0]);
    }

    const [
      { data: agData },
      { data: frmData },
      { data: vData },
      { data: bRep },
      { data: aRep },
      { data: lRep },
      { data: bStd },
      { data: lStd },
    ] = await Promise.all([
      supabase.from("agents").select("*").order("agent_name"),
      supabase.from("farmers").select("*").order("farmer_name"),
      supabase.from("vet_visits").select("*"),
      supabase
        .from("before_sale_reports")
        .select("*")
        .order("visit_date", { ascending: false }),
      supabase
        .from("after_sale_reports")
        .select("*")
        .order("visit_date", { ascending: false }),
      supabase
        .from("layer_reports")
        .select("*")
        .order("visit_date", { ascending: false }),
      supabase.from("broiler_standards").select("*"),
      supabase.from("layer_standards").select("*"),
    ]);

    // Hoisted so the cache save at the end of this load can use it.
    let sMap: Record<number, any> = {};
    try {
      const { data: sStd } = await supabase
        .from("sonali_standards")
        .select("*");
      (sStd || []).forEach((s: any) => {
        if (s.age_days !== undefined) sMap[Number(s.age_days)] = s;
      });
      setSonaliStds(sMap);
    } catch {
      sMap = {};
      setSonaliStds({});
    }

    // Sheds + sub-agents: try/catch so the page keeps working
    // pre-migration (missing table/column) — getShedsForFarmer then
    // falls back to the legacy synthesised shed.
    let shedRows: any[] = [];
    try {
      const { data: shData } = await supabase
        .from("farmer_sheds")
        .select("*");
      shedRows = shData || [];
    } catch {
      shedRows = [];
    }

    let subAgentRows: any[] = [];
    try {
      const { data: saData } = await supabase
        .from("sub_agents")
        .select("id, sub_agent_name, agent_id");
      subAgentRows = saData || [];
    } catch {
      subAgentRows = [];
    }
    setSubAgents(subAgentRows);

    // Companies (for the Company filter): try/catch so the page keeps
    // working if the table is unreachable — the filter then passes all rows.
    let companyRows: any[] = [];
    try {
      const { data: coData } = await supabase
        .from("companies")
        .select("id, company_name")
        .order("company_name");
      companyRows = coData || [];
    } catch {
      companyRows = [];
    }
    setCompanies(companyRows);

    // Users list (for the User filter): try/catch so the page keeps
    // working if the query fails — the filter then simply has no users.
    let userRows: any[] = [];
    try {
      const { data: usData } = await supabase
        .from("users")
        .select("id, full_name, upazilas")
        .order("full_name");
      userRows = usData || [];
    } catch {
      userRows = [];
    }
    // An empty list is never trusted on its own: the cache-first read can
    // serve a stale empty IndexedDB copy (cached when the first fetch
    // failed), and the page cache would then freeze the filter at 0.
    // Retry once straight from the network. If that also returns nothing,
    // the block is Supabase-side — check RLS on the users table.
    if (userRows.length === 0 && navigator.onLine) {
      try {
        const { data: netData } = await rawSupabase
          .from("users")
          .select("id, full_name, upazilas")
          .order("full_name");
        if (Array.isArray(netData) && netData.length > 0) userRows = netData;
      } catch {
        // Keep the cached (empty) result — offline or unreachable.
      }
    }
    setUsersList(userRows);

    const allFarmers = (frmData || []).map((f: any) => ({
      ...f,
      farmer_sheds: shedRows.filter(
        (s: any) => String(s.farmer_id) === String(f.id)
      ),
    }));
    const allowedFarmers = fullAccess
      ? allFarmers
      : allFarmers.filter((f: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(f.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(f.district);
          return matchUpa && matchDist;
        });

    const allAgents = agData || [];
    const allowedAgents = fullAccess
      ? allAgents
      : allAgents.filter((a: any) => {
          const aUpa = a.upazila || a.upazila_name || "";
          const matchUpa =
            uUpazilas.length === 0 || !aUpa || uUpazilas.includes(aUpa);
          const hasAllowedFarmer = allowedFarmers.some(
            (f: any) => String(f.agent_id) === String(a.id)
          );
          return matchUpa || hasAllowedFarmer;
        });

    setAgents(allowedAgents);
    setFarmers(allowedFarmers);

    const vMap: Record<string, any> = {};
    (vData || []).forEach((v: any) => (vMap[String(v.id)] = v));
    setVisitsMap(vMap);

    setBeforeReports(bRep || []);
    setAfterReports(aRep || []);
    setLayerReports(lRep || []);

    const bMap: Record<number, any> = {};
    (bStd || []).forEach((s: any) => (bMap[Number(s.age_days)] = s));
    setBroilerStds(bMap);

    const lMap: Record<string, any> = {};
    (lStd || []).forEach((s: any) => {
      const w = Number(s.age_weeks);
      const breedRaw = String(s.breed || "shaver star cross 579 brown")
        .trim()
        .toLowerCase();
      lMap[`${breedRaw}_${w}`] = s;
      if (breedRaw.includes("white")) {
        lMap[`white_${w}`] = s;
      } else {
        lMap[`brown_${w}`] = s;
      }
      if (!lMap[`any_${w}`]) {
        lMap[`any_${w}`] = s;
      }
    });
    setLayerStds(lMap);

    // Save the page state so returning to this page shows data instantly
    // with no loading spinner. Access inputs are cached alongside the data
    // (this page has no assignedCategories gate — access is role +
    // upazilas/districts only; the users query selects just those three).
    // savedUpazilaFilter mirrors the single-upazila auto-select above so
    // the restored filter matches what the user actually saw.
    const savedUpazilaFilter =
      !fullAccess && uUpazilas.length === 1 ? uUpazilas[0] : upazilaFilter;
    cachedPageState = {
      initialized: true,
      userRole: role,
      assignedUpazilas: uUpazilas,
      assignedDistricts: uDistricts,
      agents: allowedAgents,
      farmers: allowedFarmers,
      subAgents: subAgentRows,
      visitsMap: vMap,
      beforeReports: bRep || [],
      afterReports: aRep || [],
      layerReports: lRep || [],
      broilerStds: bMap,
      sonaliStds: sMap,
      layerStds: lMap,
      companies: companyRows,
      usersList: userRows,
      currentUserId: uid,
      userProfile: uProfile,
      companyFilter,
      selectedCategory,
      reportType,
      userFilter,
      upazilaFilter: savedUpazilaFilter,
      agentFilter,
      farmerFilter,
      feedTypeFilter,
      isMultiAge,
      ageFilter,
      ageFrom,
      ageTo,
      monthFilter,
      isMultiDate,
      fromDate,
      toDate,
    };

    setLoading(false);
  }

  function getMeatStd(catName?: string, ageDays?: number) {
    const d = Number(ageDays || 0);
    const c = String(catName || "Broiler")
      .trim()
      .toLowerCase();
    if (c === "sonali") return sonaliStds[d] || null;
    return broilerStds[d] || null;
  }

  function getLayerStd(breedName?: string, ageWks?: number) {
    const w = Math.min(100, Math.max(1, Number(ageWks || 1)));
    const bLower = String(breedName || "")
      .trim()
      .toLowerCase();
    if (layerStds[`${bLower}_${w}`]) return layerStds[`${bLower}_${w}`];
    const isWhite =
      bLower.includes("white") ||
      bLower.includes("decalb") ||
      bLower.includes("bb 300");
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
      layerStds[`any_${w}`] ||
      null
    );
  }

  // ---------- WhatsApp message builders (copied from VisitsPage,
  // adapted to the enriched report row: r.visit is the visit, r is the
  // report, r.farmer / r.agent / r.subAgentName / r.shedDisplayName
  // carry the rest). They build the message and open the popup below;
  // they never open wa.me directly. ----------
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
  // farmers show the agent name as before.
  function getWhatsAppAgentLine(
    agent: any,
    subAgentName?: string
  ): string {
    const subName = String(subAgentName || "").trim();
    return subName ? `sub-agent : ${subName}` : agent?.agent_name || "-";
  }

  function generateBroilerWhatsAppText(r: any) {
    const visit = r.visit || {};
    const rep = r;
    const farmer = r.farmer || {};
    const agent = r.agent || {};
    const vDate = r.visit_date || "";
    const catName = r.category || visit.category || "Broiler";
    const ageDays = calcAgeDays(visit.hatch_date, vDate);
    const std = getMeatStd(catName, ageDays);
    const hasStd = Boolean(std && Number(std.weight) > 0);

    // After-sale rows store weight / mortality under different field
    // names, and feed per bird is derived — fall back the same way the
    // table does so the message shows real values for both kinds.
    const tBirds = Number(visit.t_birds || 0);
    const feedGmFallback =
      tBirds > 0
        ? ((Number(rep.feed_consumption_kg || 0) * 1000) / tBirds).toFixed(1)
        : 0;
    const actualWt = Number(rep.weight_gm_b ?? rep.weight ?? 0);
    const stdWt = Number(std?.weight || 0);
    const wtDiff = actualWt - stdWt;
    const wtDiffText =
      hasStd && stdWt > 0
        ? `${Math.abs(wtDiff).toFixed(0)}gm ${wtDiff >= 0 ? "more" : "less"}`
        : "";

    const actualFeed = Number(rep.feed_gm_b ?? feedGmFallback ?? 0);
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

    const actualFarmerName = r.shedDisplayName || visit.batch_label || "-";

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

Agent: ${getWhatsAppAgentLine(agent, r.subAgentName)}
Farmer Name - ${actualFarmerName}
Farm Address - ${[farmer.village, farmer.upazila].filter(Boolean).join(", ")}
--------------------------------------------
Hatchery - ${visit.hatchery || "-"}
Breed - ${visit.breed || "-"}
H.Date: ${formatDmy(visit.hatch_date)}
--------------------------------------------
Total Bird (Pcs) - ${visit.t_birds || 0}
Age (Day) - ${ageDays} days
Mortality (Pcs) - ${rep.mortality_pcs ?? rep.mortality ?? 0} pcs

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

  // Layer WhatsApp Message: Disease line included, no Mortality line.
  function generateLayerWhatsAppText(r: any) {
    const visit = r.visit || {};
    const rep = r;
    const farmer = r.farmer || {};
    const agent = r.agent || {};
    const vDate = rep.visit_date || "";
    const ageWks = Math.max(
      1,
      Math.floor(calcAgeDays(visit.hatch_date, vDate) / 7)
    );

    const breedName = visit.breed || rep.breed || "Shaver Star Cross 579 Brown";
    const std = getLayerStd(breedName, ageWks);

    const hatcheryName = visit.hatchery || rep.hatchery || "Nourish";
    const feedCompName = rep.feed_company || "Nourish";

    const headerTitle = buildWhatsAppHeaderLine(
      hatcheryName,
      feedCompName,
      "Layer",
      "Layer farm visit report"
    );

    const actualFarmerName = r.shedDisplayName || visit.batch_label || "-";
    const displayBirds = Number(visit.t_birds || 0);

    const rmkText = getRemarkText(rep);
    const remarksLine = rmkText ? `\nRemarks: ${rmkText}\n` : "";

    const msg = `Assalamu Alaikum sir.
${headerTitle}
Visit Date: ${formatDmy(vDate)}

Agent Name: ${getWhatsAppAgentLine(agent, r.subAgentName)}
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

  function handleMonthSelect(ym: string) {
    setMonthFilter(ym);
    if (!ym) {
      setFromDate("");
      setToDate("");
      return;
    }
    const [y, m] = ym.split("-").map(Number);
    const lastDay = new Date(y, m, 0).getDate();
    setIsMultiDate(true);
    setFromDate(`${ym}-01`);
    setToDate(`${ym}-${String(lastDay).padStart(2, "0")}`);
  }

  // Company filter data: name lookup by company id, whether a company
  // named Nourish exists, and the dropdown list (all companies except
  // Nourish, sorted by name — Nourish sits first as its own option).
  const companyNameById: Record<string, string> = {};
  companies.forEach((c: any) => {
    companyNameById[String(c.id)] = String(c.company_name || "");
  });
  const nourishExists = companies.some(
    (c: any) =>
      String(c.company_name || "").trim().toLowerCase() === "nourish"
  );
  const otherCompanies = companies
    .filter(
      (c: any) =>
        String(c.company_name || "").trim().toLowerCase() !== "nourish"
    )
    .sort((a: any, b: any) =>
      String(a.company_name || "").localeCompare(
        String(b.company_name || "")
      )
    );

  // Upazila options strictly scoped to user's assigned upazilas
  const upazilaOptions = Array.from(
    new Set(
      (!isFullAccess && assignedUpazilas.length > 0
        ? assignedUpazilas
        : [
            ...farmers.map((f) => f.upazila),
            ...agents.map((a) => a.upazila || a.upazila_name),
          ]
      ).filter(Boolean)
    )
  ).sort();

  const availableAgents = agents.filter((a) => {
    if (upazilaFilter === "all") return true;
    const aUpa = a.upazila || a.upazila_name || "";
    const hasFarmerInUpa = farmers.some(
      (f) =>
        String(f.agent_id) === String(a.id) && f.upazila === upazilaFilter
    );
    return aUpa === upazilaFilter || hasFarmerInUpa;
  });

  const availableFarmers = farmers.filter((f) => {
    if (upazilaFilter !== "all" && f.upazila !== upazilaFilter) return false;
    if (agentFilter !== "all" && String(f.agent_id) !== String(agentFilter))
      return false;
    return true;
  });

  // Users visible in the User filter: full-access (manager / super admin)
  // sees every user; an officer sees users sharing at least one upazila
  // with their own assigned upazilas — including themselves. An officer
  // with no assigned upazilas is unrestricted, same as the area scoping.
  const visibleUsers = usersList.filter((u: any) => {
    if (isFullAccess || assignedUpazilas.length === 0) return true;
    if (String(u.id) === String(currentUserId)) return true;
    const uUpa = parseList(u.upazilas);
    return uUpa.some((x: string) => assignedUpazilas.includes(x));
  });

  // Combine raw reports based on selectedCategory and reportType
  const rawCombinedReports: any[] = [];
  if (selectedCategory === "all") {
    if (reportType === "all" || reportType === "before_sale") {
      beforeReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "before_sale" })
      );
    }
    if (reportType === "all" || reportType === "after_sale") {
      afterReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "after_sale" })
      );
    }
    if (reportType === "all") {
      layerReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "layer" })
      );
    }
  } else if (selectedCategory === "Layer") {
    layerReports.forEach((r) =>
      rawCombinedReports.push({ ...r, _repKind: "layer" })
    );
  } else {
    if (reportType === "all" || reportType === "before_sale") {
      beforeReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "before_sale" })
      );
    }
    if (reportType === "all" || reportType === "after_sale") {
      afterReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "after_sale" })
      );
    }
  }

  const allowedFarmerIds = new Set(farmers.map((f) => String(f.id)));

  // Batch numbers count per SHED, not per farmer (same rule as
  // VisitsPage): each shed's batches start again at 1. A visit without
  // a shed_id on a multi-shed farmer counts in its own "none" group.
  const visitSeqMap: Record<string, number> = {};
  {
    const visitsByFarmer: Record<string, any[]> = {};
    Object.values(visitsMap).forEach((v: any) => {
      const k = String(v.farmer_id);
      (visitsByFarmer[k] = visitsByFarmer[k] || []).push(v);
    });
    Object.values(visitsByFarmer).forEach((list: any[]) => {
      const sorted = [...list].sort((a: any, b: any) =>
        String(a.created_at || "").localeCompare(String(b.created_at || ""))
      );
      const perShedCount: Record<string, number> = {};
      sorted.forEach((v: any) => {
        const vFarmer =
          farmers.find((f: any) => String(f.id) === String(v.farmer_id)) ||
          {};
        const vShed = getShedForVisit(v, vFarmer);
        const shedKey = String(vShed?.id || "none");
        perShedCount[shedKey] = (perShedCount[shedKey] || 0) + 1;
        visitSeqMap[String(v.id)] = perShedCount[shedKey];
      });
    });
  }

  // Enrich and Filter Reports
  const filteredReportsBase = rawCombinedReports
    .map((r) => {
      const visit = visitsMap[String(r.vet_visit_id)] || {};
      const farmer =
        farmers.find((f) => String(f.id) === String(r.farmer_id)) || {};
      const agent =
        agents.find(
          (a) => String(a.id) === String(r.agent_id || farmer.agent_id)
        ) || {};
      const cat =
        r._repKind === "layer"
          ? "Layer"
          : visit.category || farmer.farm_type || "Broiler";
      const upazila =
        farmer.upazila || agent.upazila || agent.upazila_name || "";
      const hatchery = r.hatchery || visit.hatchery || "-";
      const breed = r.breed || visit.breed || "-";
      const sheds = getShedsForFarmer(farmer);
      const shed = getShedForVisit(visit, farmer);
      const shedIndex = shed ? sheds.indexOf(shed) : 0;
      const shedShort = shed
        ? getShedShortName(farmer, shed, shedIndex, sheds.length)
        : "";
      const shedDisplayName = shed
        ? getShedDisplayName(farmer, shed, shedIndex, sheds.length)
        : String(farmer?.farmer_name || "Unknown Farmer");
      const subAgent = subAgents.find(
        (sa: any) => String(sa.id) === String(farmer?.sub_agent_id)
      );
      const subAgentName = subAgent?.sub_agent_name || "";
      const seq = visitSeqMap[String(r.vet_visit_id)] || 1;
      const batchLabel = visit.batch_label || `Batch ${seq}`;
      // Age exactly as the table shows it: days for meat ("28d"),
      // weeks for Layer ("8w"). ageNum is the number shown, ageDays the
      // underlying days (used only to sort the dropdown chronologically).
      const ageDays = calcAgeDays(visit?.hatch_date, r.visit_date);
      const ageWks = Math.max(1, Math.floor(ageDays / 7));
      const ageLabel = r._repKind === "layer" ? `${ageWks}w` : `${ageDays}d`;
      const ageNum = r._repKind === "layer" ? ageWks : ageDays;

      return {
        ...r,
        visit,
        farmer,
        agent,
        upazila,
        hatchery,
        breed,
        category: cat,
        shed,
        shedShort,
        shedDisplayName,
        subAgentName,
        batchSeq: seq,
        batchLabel,
        ageDays,
        ageLabel,
        ageNum,
      };
    })
    .filter((item) => {
      if (!isFullAccess && assignedUpazilas.length > 0) {
        const matchAllowed =
          allowedFarmerIds.has(String(item.farmer_id)) ||
          assignedUpazilas.includes(item.upazila);
        if (!matchAllowed) return false;
      }
      if (
        selectedCategory !== "all" &&
        item.category?.toLowerCase() !== selectedCategory.toLowerCase()
      ) {
        return false;
      }
      if (upazilaFilter !== "all" && item.upazila !== upazilaFilter) {
        return false;
      }
      if (
        agentFilter !== "all" &&
        String(item.agent_id || item.farmer?.agent_id) !== String(agentFilter)
      ) {
        return false;
      }
      if (
        farmerFilter !== "all" &&
        String(item.farmer_id) !== String(farmerFilter)
      ) {
        return false;
      }
      // User filter: keeps only rows this user added. AND-ed with the
      // area scoping above, so it never shows reports outside the
      // viewer's own assigned area.
      if (
        userFilter !== "all" &&
        String(item.created_by || "") !== String(userFilter)
      ) {
        return false;
      }
      // Company filter: "Nourish" by default; "__others__" = every company
      // except Nourish; otherwise a single company name. If no company
      // named Nourish exists in the loaded list, the filter is skipped so
      // the table never goes empty because of it.
      if (companyFilter === "__others__") {
        const nm = (
          companyNameById[String(item.farmer?.company_id || "")] || ""
        )
          .trim()
          .toLowerCase();
        if (!nm || nm === "nourish") return false;
      } else if (nourishExists) {
        const nm = (
          companyNameById[String(item.farmer?.company_id || "")] || ""
        )
          .trim()
          .toLowerCase();
        if (nm !== companyFilter.trim().toLowerCase()) return false;
      }
      const vDate = String(item.visit_date || "").slice(0, 10);
      if (fromDate && vDate < fromDate) return false;
      if (toDate && vDate > toDate) return false;
      return true;
    })
    // Newest first: visit_date desc, then created_at desc (latest report
    // of the same date on top), then id desc as a final deterministic
    // tiebreak for rows with equal or missing created_at.
    .sort((a, b) => {
      const d = String(b.visit_date || "").localeCompare(
        String(a.visit_date || "")
      );
      if (d !== 0) return d;
      const c = String(b.created_at || "").localeCompare(
        String(a.created_at || "")
      );
      if (c !== 0) return c;
      const ai = Number(a.id);
      const bi = Number(b.id);
      if (!isNaN(ai) && !isNaN(bi)) return bi - ai;
      return String(b.id ?? "").localeCompare(String(a.id ?? ""));
    });

  // Feed Type options come from the rows left AFTER all other filters,
  // so the dropdown only ever lists types actually on screen.
  const feedTypeOptions = Array.from(
    new Set(
      filteredReportsBase
        .map((r: any) => String(r.feed_type || "").trim())
        .filter(Boolean)
    )
  ).sort();

  const afterFeedType =
    feedTypeFilter === "all"
      ? filteredReportsBase
      : filteredReportsBase.filter(
          (r: any) => String(r.feed_type || "").trim() === feedTypeFilter
        );

  // Age options come from the rows left after the feed-type filter, in
  // the same labels the table shows ("28d" / "8w"), oldest-first by days.
  const ageOptions = Array.from(
    new Map(
      afterFeedType.map((r: any) => [r.ageLabel, Number(r.ageDays)])
    ).entries()
  )
    .map(([label, days]) => ({ label: label as string, days: days as number }))
    .sort((a, b) => b.days - a.days || (a.label < b.label ? -1 : 1));

  const filteredReports =
    afterFeedType.filter((r: any) => {
      if (isMultiAge) {
        const fromRaw = ageFrom.trim() === "" ? null : Number(ageFrom);
        const toRaw = ageTo.trim() === "" ? null : Number(ageTo);
        const from =
          fromRaw !== null && !isNaN(fromRaw) ? fromRaw : null;
        const to = toRaw !== null && !isNaN(toRaw) ? toRaw : null;
        if (from === null && to === null) return true;
        const n = Number(r.ageNum);
        if (from !== null && n < from) return false;
        if (to !== null && n > to) return false;
        return true;
      }
      if (ageFilter === "all") return true;
      return r.ageLabel === ageFilter;
    });

  // Egg columns only when a Layer report is actually shown: a Broiler /
  // Sonali selection has no layer rows, so the two egg columns hide.
  const showEggColumns = filteredReports.some(
    (r: any) => r._repKind === "layer"
  );

  // Same rule, other way round: the two FCR columns only when a meat
  // (Before/After Sale) report is shown — a Layer-only selection hides them.
  const showFcrColumns = filteredReports.some(
    (r: any) => r._repKind === "before_sale" || r._repKind === "after_sale"
  );

  // If another filter change leaves the picked feed type with no rows,
  // fall back to All instead of showing an empty table.
  useEffect(() => {
    if (
      feedTypeFilter !== "all" &&
      !feedTypeOptions.includes(feedTypeFilter)
    ) {
      setFeedTypeFilter("all");
    }
  }, [feedTypeOptions, feedTypeFilter]);

  // Same for the single-age dropdown.
  useEffect(() => {
    if (
      !isMultiAge &&
      ageFilter !== "all" &&
      !ageOptions.some((a) => a.label === ageFilter)
    ) {
      setAgeFilter("all");
    }
  }, [ageOptions, ageFilter, isMultiAge]);

  // Same for the company dropdown: if the picked company disappears from
  // the list, fall back to the Nourish default.
  useEffect(() => {
    if (companies.length === 0) return;
    if (companyFilter === "__others__") return;
    if (companyFilter.trim().toLowerCase() === "nourish") return;
    const exists = companies.some(
      (c: any) =>
        String(c.company_name || "").trim().toLowerCase() ===
        companyFilter.trim().toLowerCase()
    );
    if (!exists) setCompanyFilter("Nourish");
  }, [companies, companyFilter]);

  // Same for the user dropdown: if the picked user disappears from the
  // visible list, fall back to All Users.
  useEffect(() => {
    if (userFilter === "all") return;
    if (usersList.length === 0) return;
    const exists = visibleUsers.some(
      (u: any) => String(u.id) === String(userFilter)
    );
    if (!exists) setUserFilter("all");
  }, [visibleUsers, userFilter, usersList]);

  // Collect Available Months from all reports
  const allMonths = Array.from(
    new Set(
      [...beforeReports, ...afterReports, ...layerReports]
        .map((r) => String(r.visit_date || "").slice(0, 7))
        .filter((m) => m.length === 7)
    )
  )
    .sort()
    .reverse();

  // Professional Excel Worksheet Formatter
  function formatExcelSheet(
    sheet: ExcelJS.Worksheet,
    headerHex: string,
    numericCols: number[]
  ) {
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    const headerRow = sheet.getRow(1);
    headerRow.height = 26;

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: headerHex },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: false,
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFFFFFFF" } },
        left: { style: "thin", color: { argb: "FFFFFFFF" } },
        bottom: { style: "thin", color: { argb: "FFFFFFFF" } },
        right: { style: "thin", color: { argb: "FFFFFFFF" } },
      };
    });

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      row.height = 20;
      const isEven = rowNumber % 2 === 0;
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (isEven) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF8FAFC" },
          };
        }
        cell.font = { size: 10, color: { argb: "FF0F172A" } };
        cell.alignment = {
          vertical: "middle",
          horizontal: numericCols.includes(colNumber) ? "right" : "left",
          wrapText: false,
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } },
        };
      });
    });

    sheet.columns.forEach((column) => {
      let maxLen = 11;
      column.eachCell?.({ includeEmpty: true }, (cell) => {
        // Rich-text cells (Agent with the small "(sub agent)" suffix)
        // store an object — measure the joined run texts, not
        // "[object Object]".
        const rawVal = cell.value as any;
        const val = rawVal
          ? rawVal.richText
            ? rawVal.richText.map((t: any) => t.text || "").join("")
            : String(rawVal)
          : "";
        if (val.length > maxLen) maxLen = val.length;
      });
      column.width = Math.min(Math.max(maxLen + 4, 12), 38);
    });
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

      alert(`File successfully saved:\n📁 Documents/${fileName}`);

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
    alert("Failed to save file: " + (err?.message || ""));
  }
}

  // Export cell values: Farmer carries the shed inside its own cell
  // ("Raju Ahmed Shed 1"); Agent carries the sub-agent with a small
  // "(sub agent)" suffix (rich text, so only the suffix is small).
  const exportFarmerCell = (r: any) =>
    r.shedShort
      ? `${r.farmer?.farmer_name || "-"} ${r.shedShort}`
      : r.farmer?.farmer_name || "-";
  const exportAgentCell = (r: any) =>
    r.subAgentName
      ? {
          richText: [
            {
              text: r.subAgentName,
              font: { size: 10, color: { argb: "FF0F172A" } },
            },
            {
              text: " (sub agent)",
              font: { size: 8, italic: true, color: { argb: "FF64748B" } },
            },
          ],
        }
      : r.agent?.agent_name || "-";

  async function exportFilteredReportsExcel() {
    if (filteredReports.length === 0) {
      alert("No reports found to export.");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const beforeList = filteredReports.filter(
      (r) => r._repKind === "before_sale"
    );
    const afterList = filteredReports.filter(
      (r) => r._repKind === "after_sale"
    );
    const layerList = filteredReports.filter((r) => r._repKind === "layer");

    if (beforeList.length > 0) {
      const s = workbook.addWorksheet("Before_Sale_Reports");
      s.addRow([
        "Visit Date",
        "Upazila",
        "Agent",
        "Farmer",
        "Category",
        "Hatchery",
        "Breed",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt (g)",
        "Feed (Kg)",
        "Feed/B (g)",
        "Std Feed (g)",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Company",
        "Feed Type",
        "MFG",
        "Batch No.",
        "Remarks",
      ]);

      beforeList.forEach((r: any) => {
        const d = calcAgeDays(r.visit?.hatch_date, r.visit_date);
        const std = getMeatStd(r.category, d);
        s.addRow([
          r.visit_date || "-",
          r.upazila || "-",
          exportAgentCell(r),
          exportFarmerCell(r),
          r.category || "-",
          r.hatchery || "-",
          r.breed || "-",
          r.visit?.hatch_date || "-",
          Number(r.visit?.t_birds || 0),
          d,
          Number(r.mortality_pcs || 0),
          Number(r.weight_gm_b || 0),
          std?.weight || "-",
          Number(r.feed_t_kg || 0),
          Number(r.feed_gm_b || 0),
          std?.feed_intake || "-",
          r.fcr_w_out_m || "-",
          std?.fcr || "-",
          r.fcr_with_m || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          r.pro_date ? formatDmy(r.pro_date) : "-",
          r.batch_no || "-",
          getRemarkText(r) || "-",
        ]);
      });

      formatExcelSheet(
        s,
        "FF059669",
        [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]
      );
    }

    if (afterList.length > 0) {
      const s = workbook.addWorksheet("After_Sale_Reports");
      s.addRow([
        "Sale Date",
        "Upazila",
        "Agent",
        "Farmer",
        "Category",
        "Hatchery",
        "Breed",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt (g)",
        "Total Feed (Kg)",
        "Feed/B (g)",
        "Std Feed (g)",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Company",
        "Last Feed",
        "MFG",
        "Batch No.",
        "Remarks",
      ]);

      afterList.forEach((r: any) => {
        const d = calcAgeDays(r.visit?.hatch_date, r.visit_date);
        const std = getMeatStd(r.category, d);
        const tB = Number(r.visit?.t_birds || 0);
        const fGm =
          r.feed_gm_b ||
          (tB > 0
            ? ((Number(r.feed_consumption_kg || 0) * 1000) / tB).toFixed(1)
            : "-");
        s.addRow([
          r.visit_date || "-",
          r.upazila || "-",
          exportAgentCell(r),
          exportFarmerCell(r),
          r.category || "-",
          r.hatchery || "-",
          r.breed || "-",
          r.visit?.hatch_date || "-",
          tB,
          d,
          Number(r.mortality || 0),
          Number(r.weight || 0),
          std?.weight || "-",
          Number(r.feed_consumption_kg || 0),
          fGm,
          std?.feed_intake || "-",
          r.fcr_w_out_m || "-",
          std?.fcr || "-",
          r.fcr_with_m || "-",
          r.feed_mill || "-",
          r.feed || "-",
          r.pro_date ? formatDmy(r.pro_date) : "-",
          r.batch_no || "-",
          getRemarkText(r) || "-",
        ]);
      });

      formatExcelSheet(
        s,
        "FFDC2626",
        [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]
      );
    }

    if (layerList.length > 0) {
      const s = workbook.addWorksheet("Layer_Reports");
      s.addRow([
        "Visit Date",
        "Upazila",
        "Agent",
        "Farmer",
        "Category",
        "Hatchery",
        "Breed",
        "H.Date",
        "T.Birds",
        "Mortality",
        "Age (Wks)",
        "Weight (g)",
        "Std Wt (g)",
        "Feed (g/b/d)",
        "Std Feed (g)",
        "Egg Prod (%)",
        "Std Egg (%)",
        "Egg Wt (g)",
        "Std Egg Wt (g)",
        "Feed Company",
        "Feed Type",
        "MFG",
        "Batch No.",
        "Disease / Vaccine",
        "Remarks",
      ]);

      layerList.forEach((r: any) => {
        const wks = Math.max(
          1,
          Math.floor(calcAgeDays(r.visit?.hatch_date, r.visit_date) / 7)
        );
        const std = getLayerStd(r.breed, wks);
        s.addRow([
          r.visit_date || "-",
          r.upazila || "-",
          exportAgentCell(r),
          exportFarmerCell(r),
          r.category || "-",
          r.hatchery || "-",
          r.breed || "-",
          r.visit?.hatch_date || "-",
          Number(r.visit?.t_birds || 0),
          Number(r.mortality ?? 0),
          wks,
          Number(r.b_weight_g || 0),
          std?.weight || "-",
          Number(r.feed_intake_g || 0),
          std?.feed_bird_day || "-",
          Number(r.egg_production || 0),
          std?.egg_production || "-",
          Number(r.egg_weight || 0),
          std?.egg_weight || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          r.pro_date ? formatDmy(r.pro_date) : "-",
          r.batch_no || "-",
          [r.disease, r.vaccine].filter(Boolean).join(" / ") || "-",
          getRemarkText(r) || "-",
        ]);
      });

      formatExcelSheet(
        s,
        "FF4F46E5",
        [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]
      );
    }

    const today = new Date().toISOString().slice(0, 10);
    const fileName = `SmartTeam_Reports_${selectedCategory || "All"}_${today}.xlsx`;

    try {
      const excelBuffer = await workbook.xlsx.writeBuffer();
      // Shared saveBufferOrBlob helper (same as DocTracking / Farmers pages)
      await saveBufferOrBlob(excelBuffer, fileName);
    } catch (err: any) {
      console.error("Filtered reports export error:", err);
      alert("Reports export failed.");
    }
  }

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-3.5">
      {/* Filter Toggle Button — mobile only; on PC (md+) both filter sections stay open */}
      <button
        type="button"
        onClick={() => setShowFilters((o) => !o)}
        className={`inline-flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-bold shadow-sm transition-all duration-300 md:hidden ${
          showFilters
            ? "bg-slate-800 text-white"
            : "bg-emerald-600 hover:bg-emerald-700 text-white"
        }`}
      >
        <Filter size={13} />
        <span>Filter</span>
        <ChevronDown
          size={14}
          className={`transition-transform duration-300 ${
            showFilters ? "rotate-180" : "rotate-0"
          }`}
        />
      </button>

      {/* Top Bar + Filter Bar — collapsed on mobile, forced open at md+ */}
      <div
        className={`grid transition-all duration-300 ease-in-out md:grid-rows-[1fr] md:opacity-100 md:translate-y-0 md:pointer-events-auto ${
          showFilters
            ? "grid-rows-[1fr] opacity-100 translate-y-0"
            : "grid-rows-[0fr] opacity-0 -translate-y-1 pointer-events-none"
        }`}
      >
        <div className="overflow-hidden space-y-3.5">
      {/* Top Bar: All + Category Tabs & Report Type Tabs */}
      <div className="flex flex-wrap items-center justify-between gap-2.5 bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 shadow-sm">
        <div className="flex flex-wrap items-center gap-1.5">
          {(["all", "Broiler", "Layer", "Sonali", "Fish", "Cattle"] as const).map((cat) => (
            <button
              key={cat}
              type="button"
              onClick={() => setSelectedCategory(cat)}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                selectedCategory === cat
                  ? "bg-emerald-600 text-white shadow-sm"
                  : "bg-slate-100 text-slate-700 hover:bg-slate-200"
              }`}
            >
              {cat === "all" ? "All" : cat}
            </button>
          ))}

          {selectedCategory !== "Layer" ? (
            <div className="flex items-center gap-1.5 ml-1.5 pl-2 border-l border-slate-300">
              <button
                type="button"
                onClick={() => setReportType("all")}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "all"
                    ? "bg-indigo-600 text-white"
                    : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                }`}
              >
                All
              </button>
              <button
                type="button"
                onClick={() => setReportType("before_sale")}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "before_sale"
                    ? "bg-slate-800 text-white"
                    : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                }`}
              >
                Before Sale
              </button>
              <button
                type="button"
                onClick={() => setReportType("after_sale")}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "after_sale"
                    ? "bg-red-600 text-white"
                    : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                }`}
              >
                After Sale
              </button>
            </div>
          ) : (
            <span className="ml-1.5 px-2.5 py-1 rounded-lg bg-indigo-50 border border-indigo-200 text-indigo-800 text-xs font-bold">
              Layer Performance Reports
            </span>
          )}
        </div>

        <button
          type="button"
          onClick={exportFilteredReportsExcel}
          className="bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm"
        >
          Export Excel ({filteredReports.length})
        </button>
      </div>

      {/* Filter Bar: Company, Assigned Upazila, Agent, Farmer, Month, Single/Multi Date Picker, Reset */}
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div
          className={`grid grid-cols-2 sm:grid-cols-3 ${
            isMultiAge ? "lg:grid-cols-12" : "lg:grid-cols-11"
          } gap-2 items-center text-xs`}
        >
          {/* Company Dropdown — Nourish first (default), then All Others, then the company list */}
          <SearchableSelect
            value={companyFilter}
            onChange={setCompanyFilter}
            placeholder="Select Company"
            options={[
              { value: "Nourish", label: "Nourish" },
              {
                value: "__others__",
                label: `All Others (${otherCompanies.length})`,
              },
              ...otherCompanies.map((c: any) => ({
                value: String(c.company_name),
                label: String(c.company_name),
              })),
            ]}
          />

          {/* Assigned Upazila Dropdown */}
          <SearchableSelect
            value={upazilaFilter}
            onChange={(v) => {
              setUpazilaFilter(v);
              setAgentFilter("all");
              setFarmerFilter("all");
            }}
            placeholder="All Upazilas"
            options={[
              { value: "all", label: "All Upazilas" },
              ...upazilaOptions.map((u) => ({ value: u, label: u })),
            ]}
          />

          {/* Agent Dropdown */}
          <SearchableSelect
            value={agentFilter}
            onChange={(v) => {
              setAgentFilter(v);
              setFarmerFilter("all");
            }}
            placeholder="All Agents"
            options={[
              {
                value: "all",
                label: `All Agents (${availableAgents.length})`,
              },
              ...availableAgents.map((a) => ({
                value: String(a.id),
                label: a.agent_name,
              })),
            ]}
          />

          {/* Farmer Dropdown */}
          <SearchableSelect
            value={farmerFilter}
            onChange={setFarmerFilter}
            placeholder="All Farmers"
            options={[
              {
                value: "all",
                label: `All Farmers (${availableFarmers.length})`,
              },
              ...availableFarmers.map((f) => ({
                value: String(f.id),
                label: `${f.farmer_name} (${f.village})`,
              })),
            ]}
          />

          {/* Feed Type Dropdown — options from the rows left after the other filters */}
          <SearchableSelect
            value={feedTypeFilter}
            onChange={setFeedTypeFilter}
            placeholder="All Feed Types"
            options={[
              {
                value: "all",
                label: `All Feed Types (${feedTypeOptions.length})`,
              },
              ...feedTypeOptions.map((t) => ({ value: t, label: t })),
            ]}
          />

          {/* Age Filter: single dropdown, or From/To range when Multi is checked */}
          <div
            className={`${
              isMultiAge ? "col-span-2 sm:col-span-2 lg:col-span-2" : ""
            } flex items-center gap-1.5 min-w-0`}
          >
            {!isMultiAge ? (
              <div className="flex-1 min-w-0">
                <SearchableSelect
                  value={ageFilter}
                  onChange={setAgeFilter}
                  placeholder="All Ages"
                  options={[
                    {
                      value: "all",
                      label: `All Ages (${ageOptions.length})`,
                    },
                    ...ageOptions.map((a) => ({
                      value: a.label,
                      label: a.label,
                    })),
                  ]}
                />
              </div>
            ) : (
              <div className="flex-1 min-w-0 grid grid-cols-2 gap-1.5">
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-bold text-slate-400 mr-1 shrink-0">
                    From:
                  </span>
                  <input
                    type="number"
                    min={0}
                    value={ageFrom}
                    onChange={(e) => setAgeFrom(e.target.value)}
                    placeholder="Age"
                    className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                  />
                </div>
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-bold text-slate-400 mr-1 shrink-0">
                    To:
                  </span>
                  <input
                    type="number"
                    min={0}
                    value={ageTo}
                    onChange={(e) => setAgeTo(e.target.value)}
                    placeholder="Age"
                    className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                  />
                </div>
              </div>
            )}

            <label
              title="Age range (From - To)"
              className="shrink-0 inline-flex items-center gap-1 px-2 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 text-[10px] font-semibold text-slate-700 cursor-pointer select-none"
            >
              <input
                type="checkbox"
                checked={isMultiAge}
                onChange={(e) => {
                  const checked = e.target.checked;
                  setIsMultiAge(checked);
                  if (checked) {
                    setAgeFilter("all");
                  } else {
                    setAgeFrom("");
                    setAgeTo("");
                  }
                }}
                className="w-3.5 h-3.5 accent-emerald-600"
              />
              <span>Multi</span>
            </label>
          </div>

          {/* Month Filter */}
          <SearchableSelect
            value={monthFilter}
            onChange={handleMonthSelect}
            placeholder="All Months"
            options={[
              { value: "", label: "All Months" },
              ...allMonths.map((m) => ({ value: m, label: `Month: ${m}` })),
            ]}
          />

          {/* Single Date OR Multi (From - To) Date Selector */}
          <div
            className={`${
              isMultiDate
                ? "col-span-2 sm:col-span-3 lg:col-span-2"
                : "col-span-2 sm:col-span-2 lg:col-span-2"
            } flex items-center gap-1.5 min-w-0`}
          >
            {!isMultiDate ? (
              <div className="flex-1 min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                <span className="text-[10px] font-bold text-slate-400 mr-1.5 shrink-0">
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
                  }}
                  className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                />
              </div>
            ) : (
              <div className="flex-1 min-w-0 grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-bold text-slate-400 mr-1 shrink-0">
                    From:
                  </span>
                  <input
                    type="date"
                    value={fromDate}
                    onChange={(e) => {
                      setFromDate(e.target.value);
                      setMonthFilter("");
                    }}
                    className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                  />
                </div>
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-bold text-slate-400 mr-1 shrink-0">
                    To:
                  </span>
                  <input
                    type="date"
                    value={toDate}
                    onChange={(e) => {
                      setToDate(e.target.value);
                      setMonthFilter("");
                    }}
                    className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                  />
                </div>
              </div>
            )}

            <label
              title="Select multiple dates (From - To)"
              className="shrink-0 inline-flex items-center gap-1 px-2 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 text-[10px] font-semibold text-slate-700 cursor-pointer select-none"
            >
              <input
                type="checkbox"
                checked={isMultiDate}
                onChange={(e) => {
                  const checked = e.target.checked;
                  setIsMultiDate(checked);
                  if (!checked) {
                    const single = fromDate || toDate || "";
                    setFromDate(single);
                    setToDate(single);
                    setMonthFilter("");
                  }
                }}
                className="w-3.5 h-3.5 accent-emerald-600"
              />
              <span>Multi</span>
            </label>
          </div>

          {/* User Dropdown — after Date, before Reset; only users sharing
              an upazila with the viewer (managers/admins see all) */}
          <SearchableSelect
            value={userFilter}
            onChange={setUserFilter}
            placeholder="All Users"
            options={[
              {
                value: "all",
                label: `All Users (${visibleUsers.length})`,
              },
              ...visibleUsers.map((u: any) => ({
                value: String(u.id),
                label: String(u.full_name || u.id),
              })),
            ]}
          />

          {/* Reset Button */}
          <button
            type="button"
            onClick={() => {
              setCompanyFilter("Nourish");
              setUpazilaFilter(
                !isFullAccess && assignedUpazilas.length === 1
                  ? assignedUpazilas[0]
                  : "all"
              );
              setAgentFilter("all");
              setFarmerFilter("all");
              setFeedTypeFilter("all");
              setUserFilter("all");
              setIsMultiAge(false);
              setAgeFilter("all");
              setAgeFrom("");
              setAgeTo("");
              setMonthFilter("");
              setIsMultiDate(false);
              setFromDate("");
              setToDate("");
            }}
            className="rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-3 py-1.5 font-semibold text-slate-700"
          >
            Reset
          </button>
        </div>
      </div>
        </div>
      </div>

      {/* Flat Reports Table (newest first) */}
      {loading ? (
        <div className="bg-white border-y sm:border border-slate-300 rounded-none sm:rounded-xl mx-[-12px] sm:mx-0 p-8 text-center text-xs text-slate-500">
          Loading reports...
        </div>
      ) : filteredReports.length === 0 ? (
        <div className="bg-white border-y sm:border border-slate-300 rounded-none sm:rounded-xl mx-[-12px] sm:mx-0 p-8 text-center text-xs text-slate-400">
          No reports found for the selected filters.
        </div>
      ) : (
        <div className="bg-white border-y sm:border border-slate-300 rounded-none sm:rounded-xl mx-[-12px] sm:mx-0 overflow-hidden shadow-sm">
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-xs whitespace-nowrap">
              <thead className="bg-slate-200">
                <tr>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Sl</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Visit Date</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Agent</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left md:sticky md:left-0 md:z-20 md:bg-slate-200">Farmer</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Category</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Type</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Hatchery</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">H.Date</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Breed</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-center">Birds</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-center">Age</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-center">Mort.</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-center">Wt / Std (g)</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-center">Feed (Kg)</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-center">Feed/B / Std (g)</th>
                  {showFcrColumns && (
                    <>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">FCR (w/o) / Std</th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">FCR (With)</th>
                    </>
                  )}
                  {showEggColumns && (
                    <>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">Egg% / Std%</th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">Egg Wt / Std (g)</th>
                    </>
                  )}
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Feed &amp; Type</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">MFG</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Batch No.</th>
                  {showEggColumns && (
                    <th className="border border-slate-300 px-2.5 py-1.5 text-left">Disease / Vaccine</th>
                  )}
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Remarks</th>
                  <th className="border border-slate-300 px-2.5 py-1.5 text-left">Actions</th>
                </tr>
              </thead>
              <tbody>
                {filteredReports.map((r: any, idx: number) => {
                  const isLayer = r._repKind === "layer";
                  const isBefore = r._repKind === "before_sale";
                  const d = calcAgeDays(r.visit?.hatch_date, r.visit_date);
                  const wks = Math.max(1, Math.floor(d / 7));
                  const std = isLayer
                    ? getLayerStd(r.breed, wks)
                    : getMeatStd(r.category, d);
                  const tB = Number(r.visit?.t_birds || 0);
                  const fGm =
                    r.feed_gm_b ||
                    (tB > 0
                      ? (
                          (Number(r.feed_consumption_kg || 0) * 1000) /
                          tB
                        ).toFixed(1)
                      : "-");
                  return (
                    <tr
                      key={`${r._repKind}_${r.id ?? idx}`}
                      onClick={() =>
                        setSelectedReportKey((prev) =>
                          prev === `${r._repKind}_${r.id ?? idx}`
                            ? null
                            : `${r._repKind}_${r.id ?? idx}`
                        )
                      }
                      className={`${
                        selectedReportKey ===
                        `${r._repKind}_${r.id ?? idx}`
                          ? "bg-emerald-200 hover:bg-emerald-200"
                          : idx % 2 === 1
                            ? "bg-slate-100 hover:bg-emerald-100"
                            : "bg-white hover:bg-emerald-100"
                      } cursor-pointer leading-tight`}
                    >
                      <td className="border border-slate-300 px-2.5 py-1 text-center">
                        {idx + 1}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {formatDmy(r.visit_date)}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {r.subAgentName ? (
                          <>
                            {r.subAgentName}{" "}
                            <span className="text-[10px] text-slate-400">
                              sub-ag
                            </span>
                          </>
                        ) : (
                          r.agent?.agent_name || "-"
                        )}
                      </td>
                      <td
                        className="border border-slate-300 px-2.5 py-1 font-semibold md:sticky md:left-0 md:z-10 md:bg-inherit"
                        title={r.shedDisplayName}
                      >
                        <div>{r.farmer?.farmer_name || "-"}</div>
                        {r.shedShort && (
                          <div className="text-[10px] font-normal text-slate-500 leading-tight">
                            {r.shedShort}
                          </div>
                        )}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {r.category}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {isLayer
                          ? "Layer"
                          : isBefore
                          ? "Before Sale"
                          : "After Sale"}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {r.hatchery || "-"}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {formatDmy(r.visit?.hatch_date)}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {r.breed || "-"}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 text-center">
                        {tB}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 text-center font-bold">
                        {isLayer ? `${wks}w` : `${d}d`}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 text-center">
                        {isLayer
                          ? r.mortality ?? 0
                          : isBefore
                          ? r.mortality_pcs
                          : r.mortality}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 text-center">
                        {isLayer ? (
                          <>
                            <b>{r.b_weight_g}</b> / {std?.weight || "-"}
                          </>
                        ) : isBefore ? (
                          <>
                            <b>{r.weight_gm_b}</b> / {std?.weight || "-"}
                          </>
                        ) : (
                          <>
                            <b>{r.weight}</b> / {std?.weight || "-"}
                          </>
                        )}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 text-center">
                        {isLayer
                          ? r.feed_total_kg
                          : isBefore
                          ? r.feed_t_kg
                          : r.feed_consumption_kg}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1 text-center">
                        {isLayer ? (
                          <>
                            <b>{r.feed_intake_g}</b> /{" "}
                            {std?.feed_bird_day || "-"}
                          </>
                        ) : isBefore ? (
                          <>
                            <b>{r.feed_gm_b}</b> / {std?.feed_intake || "-"}
                          </>
                        ) : (
                          <>
                            <b>{fGm}</b> / {std?.feed_intake || "-"}
                          </>
                        )}
                      </td>
                      {showFcrColumns && (
                        <>
                          <td className="border border-slate-300 px-2.5 py-1 text-center font-bold">
                            {isLayer ? "-" : `${r.fcr_w_out_m} / ${std?.fcr || "-"}`}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1 text-center">
                            {isLayer ? "-" : r.fcr_with_m}
                          </td>
                        </>
                      )}
                      {showEggColumns && (
                        <>
                          <td className="border border-slate-300 px-2.5 py-1 text-center font-bold text-indigo-700">
                            {isLayer
                              ? `${r.egg_production}% / ${std?.egg_production || "-"}%`
                              : "-"}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1 text-center">
                            {isLayer ? (
                              <>
                                <b>{r.egg_weight}</b> / {std?.egg_weight || "-"}
                              </>
                            ) : (
                              "-"
                            )}
                          </td>
                        </>
                      )}
                      <td className="border border-slate-300 px-2.5 py-1">
                        {isBefore || isLayer ? (
                          <>
                            {r.feed_company || "-"}{" "}
                            {r.feed_type ? `(${r.feed_type})` : ""}
                          </>
                        ) : (
                          <>
                            {r.feed_mill || "-"}{" "}
                            {r.feed ? `(Last Feed: ${r.feed})` : ""}
                          </>
                        )}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {r.pro_date ? formatDmy(r.pro_date) : "-"}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        {r.batch_no || "-"}
                      </td>
                      {showEggColumns && (
                        <td className="border border-slate-300 px-2.5 py-1">
                          {isLayer
                            ? [r.disease, r.vaccine]
                                .filter(Boolean)
                                .join(" / ") || "-"
                            : "-"}
                        </td>
                      )}
                      <td className="border border-slate-300 px-2.5 py-1">
                        {getRemarkText(r) || "-"}
                      </td>
                      <td className="border border-slate-300 px-2.5 py-1">
                        <div className="flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() =>
                              navigate(
                                r.visit?.id
                                  ? `/visits?focus=${r.visit.id}&focusFarmer=${r.farmer_id}&focusReport=${r.id}&focusKind=${r._repKind}`
                                  : `/visits?focusFarmer=${r.farmer_id}`
                              )
                            }
                            title="Visit"
                            aria-label="Visit"
                            className="bg-violet-600 hover:bg-violet-700 text-white px-2 py-0.5 rounded text-[11px] font-semibold inline-flex items-center gap-0.5"
                          >
                            <Stethoscope size={11} /> Visit
                          </button>
                          {(r.created_by === currentUserId ||
                            r.created_by == null) && (
                            <button
                              type="button"
                              onClick={() =>
                                isLayer
                                  ? generateLayerWhatsAppText(r)
                                  : generateBroilerWhatsAppText(r)
                              }
                              className="bg-green-600 hover:bg-green-700 text-white px-2 py-0.5 rounded text-[11px] font-semibold whitespace-nowrap"
                            >
                              WhatsApp Msg
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
      {/* WhatsApp Popup (same as VisitsPage) */}
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