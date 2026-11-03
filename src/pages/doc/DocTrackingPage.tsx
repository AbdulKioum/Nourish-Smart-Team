import React, { useEffect, useState, useRef } from "react";
import XLSX from "xlsx-js-style";
import { supabase } from "../../lib/supabase";
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
} from "react-icons/fa";
import { RotateCcw } from "lucide-react";

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
  Broiler: ["Arber Acress", "Ross", "Cobb 500", "Indian River", "Hubbard"],
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

// PCD / DOC Type to Category & Breed Auto-Mapper
function mapPcdToCategoryAndBreed(pcd: string): {
  category: string;
  breed: string;
} {
  const code = String(pcd || "").trim().toUpperCase();
  if (code.startsWith("AA"))
    return { category: "Broiler", breed: "Arber Acress" };
  if (code.startsWith("HE")) return { category: "Broiler", breed: "Hubbard" };
  if (code.startsWith("RO")) return { category: "Broiler", breed: "Ross" };

  if (code.startsWith("SA")) return { category: "Sonali", breed: "Sasso" };
  if (code.startsWith("BP"))
    return { category: "Sonali", breed: "Bornali Plus" };
  if (code.startsWith("BR")) return { category: "Sonali", breed: "Bornali" };
  if (code.startsWith("CB")) return { category: "Sonali", breed: "Brown Cock" };
  if (code.startsWith("CW")) return { category: "Sonali", breed: "White Cock" };

  if (code.startsWith("SH"))
    return { category: "Layer", breed: "Shaver Star Cross 579 Brown" };
  if (code.startsWith("BO"))
    return { category: "Layer", breed: "Bovans White" };

  return { category: "Broiler", breed: "Arber Acress" };
}

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
  feed_company: string;
  remark: string;
  category?: string;
  breed?: string;
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
} | null = null;

export default function DocTrackingPage() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [loading, setLoading] = useState(!cachedPageState?.initialized);
  const [importing, setImporting] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");
  const [successMsg, setSuccessMsg] = useState("");

  // Current User Info
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

  // Data States
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

  // Compact Filters
  const todayStr = new Date().toISOString().split("T")[0];
  const [fromDate, setFromDate] = useState<string>(
    cachedPageState?.fromDate || ""
  );
  const [toDate, setToDate] = useState<string>(cachedPageState?.toDate || "");
  const [monthFilter, setMonthFilter] = useState<string>(
    cachedPageState?.monthFilter || ""
  );
  const [pcdFilter, setPcdFilter] = useState<string>("");
  const [upazilaFilter, setUpazilaFilter] = useState<string>("");
  const [officerFilter, setOfficerFilter] = useState<string>("");
  const [searchAgent, setSearchAgent] = useState<string>("");

  // Add / Edit Tracking Modal State
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
    category: "Broiler",
    breed: "Arber Acress",
    feed_company: "",
    remark: "",
  });

  // Export Modal State
  const [isExportModalOpen, setIsExportModalOpen] = useState(false);
  const [exportSelectedTypes, setExportSelectedTypes] = useState<string[]>([]);

  // Tracked Report Modal State
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
          "id, agent_id, company_id, farmer_name, mobile, village, union_name, upazila, district, companies(company_name)"
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

    const loadedAgents: DocAgent[] = agentData || [];
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

    const loadedSupplies: DocSupply[] = supData || [];
    const loadedTrackings: DocTracking[] = (trkData || []).map((t: any) => {
      let trackerName = t.tracked_by_name;
      if (!trackerName && t.created_by) {
        const foundOff = parsedOfficers.find((o) => o.id === t.created_by);
        if (foundOff) trackerName = foundOff.full_name;
      }
      return {
        ...t,
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
    };
  }

  // ==================================================
  // EXCEL HELPERS
  // ==================================================
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

        const rawPcd = String(getVal(["pcd", "doc type"]) || "")
          .replace(/<[^>]*>/g, "")
          .trim();
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
        setMonthFilter("");
        if (cachedPageState) {
          cachedPageState.fromDate = importedDate;
          cachedPageState.toDate = importedDate;
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

  // ==================================================
  // COLLABORATED FEED AGENT & FARMERS LOOKUP
  // ==================================================
  function getCollaboratedFeedAgentId(
    target: typeof activeTarget
  ): number | null {
    if (!target) return null;
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
    return inUpazila ? inUpazila.id : null;
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

  // ==================================================
  // ADD, EDIT & DELETE FARMER TRACKING (WITH VET VISIT SYNC)
  // ==================================================
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

  function openAddTrackingModal(params: {
    doc_id: string;
    agent_name: string;
    upazila: string;
    district: string;
    feed_agent_id?: number | null;
    pcd: string;
    h_date: string;
    supply_id?: string;
    remainingQty: number;
  }) {
    if (params.remainingQty <= 0) {
      alert("all DOC are tracked");
      return;
    }

    const defaultFeedComp =
      companyNames.find((c) => c.toLowerCase().includes("nourish")) ||
      companyNames[0] ||
      "";

    const { category: autoCat, breed: autoBreed } = mapPcdToCategoryAndBreed(
      params.pcd
    );

    setEditingTracking(null);
    setActiveTarget(params);
    setSelectedFarmerId("");
    setTrackForm({
      farmer_name: "",
      address: "",
      mobile: "",
      qty: String(params.remainingQty),
      category: autoCat,
      breed: autoBreed,
      feed_company: defaultFeedComp,
      remark: "",
    });
    setIsTrackModalOpen(true);
  }

  function openEditTrackingModal(
    tr: DocTracking,
    agent: DocAgent,
    currentUntrackedQty: number
  ) {
    if (!canModifyTracking(tr)) {
      alert("You can only edit tracking entries added by you.");
      return;
    }

    const { category: autoCat, breed: autoBreed } = mapPcdToCategoryAndBreed(
      tr.pcd
    );

    setEditingTracking(tr);
    setActiveTarget({
      doc_id: agent.doc_id,
      agent_name: agent.agent_name,
      upazila: agent.upazila,
      district: agent.district,
      feed_agent_id: agent.feed_agent_id,
      pcd: tr.pcd,
      h_date: tr.h_date,
      supply_id: tr.supply_id || undefined,
      remainingQty: currentUntrackedQty + Number(tr.qty || 0),
    });
    setSelectedFarmerId(tr.farmer_id ? String(tr.farmer_id) : "");
    setTrackForm({
      farmer_name: tr.farmer_name || "",
      address: tr.address || "",
      mobile: tr.mobile || "",
      qty: String(tr.qty || ""),
      category: tr.category || autoCat,
      breed: tr.breed || autoBreed,
      feed_company: tr.feed_company || companyNames[0] || "",
      remark: tr.remark || "",
    });
    setIsTrackModalOpen(true);
  }

  function handleSelectFarmerFromDropdown(farmerId: string) {
    setSelectedFarmerId(farmerId);
    if (!farmerId) return;

    const f = collaboratedFarmers.find((item) => String(item.id) === farmerId);
    if (!f) return;

    const compName = f.companies?.company_name || "";
    const matchedFeedComp =
      companyNames.find((c) => c.toLowerCase() === compName.toLowerCase()) ||
      trackForm.feed_company ||
      companyNames[0] ||
      "";

    const addr = [f.village, f.union_name, f.upazila]
      .filter(Boolean)
      .join(", ");

    setTrackForm((prev) => ({
      ...prev,
      farmer_name: f.farmer_name || "",
      address: addr,
      mobile: f.mobile || "",
      feed_company: matchedFeedComp,
    }));
  }

  async function handleSaveTracking(e: React.FormEvent) {
    e.preventDefault();
    if (!activeTarget) return;

    const numQty = Number(trackForm.qty);
    if (!trackForm.farmer_name.trim() || numQty <= 0) {
      setErrorMsg("Farmer Name and valid DOC Qty are required.");
      return;
    }

    if (numQty > activeTarget.remainingQty) {
      alert(
        `DOC Qty (${numQty}) cannot exceed remaining untracked DOC (${activeTarget.remainingQty}).`
      );
      return;
    }

    setSubmitting(true);
    setErrorMsg("");

    try {
      const finalCategory = trackForm.category || "Broiler";
      const finalBreed = trackForm.breed || "Arber Acress";

      if (editingTracking) {
        if (!canModifyTracking(editingTracking)) {
          throw new Error("Permission denied to edit this tracking.");
        }

        const updateData: Record<string, any> = {
          farmer_name: trackForm.farmer_name.trim(),
          address: trackForm.address.trim(),
          mobile: trackForm.mobile.trim(),
          qty: numQty,
          feed_company: trackForm.feed_company,
          remark: trackForm.remark.trim(),
          category: finalCategory,
          breed: finalBreed,
        };

        const { error } = await supabase
          .from("doc_trackings")
          .update(updateData)
          .eq("id", editingTracking.id);

        if (error) {
          delete updateData.category;
          delete updateData.breed;
          const { error: retryErr } = await supabase
            .from("doc_trackings")
            .update(updateData)
            .eq("id", editingTracking.id);
          if (retryErr) throw retryErr;
        }
      } else {
        // 1. Sync Farmer & Vet Visit (Direct Collaboration with Farmer Status & Vet Visits)
        const targetFeedAgentId = getCollaboratedFeedAgentId(activeTarget);
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
        let targetFarmerId = selectedFarmerId ? Number(selectedFarmerId) : null;

        if (!targetFarmerId) {
          const existingFarmer = farmersList.find(
            (f) =>
              f.farmer_name?.trim().toLowerCase() ===
                trackForm.farmer_name.trim().toLowerCase() &&
              (!targetFeedAgentId ||
                Number(f.agent_id) === Number(targetFeedAgentId))
          );

          if (existingFarmer) {
            targetFarmerId = existingFarmer.id;
          } else if (targetFeedAgentId) {
            const newFarmerPayload: Record<string, any> = {
              company_id: targetCompId,
              agent_id: targetFeedAgentId,
              farmer_name: trackForm.farmer_name.trim(),
              mobile: trackForm.mobile.trim(),
              village: trackForm.address.trim(),
              upazila: activeTarget.upazila,
              district: activeTarget.district || "Mymensingh",
              farm_category: "Poultry",
              farm_type: finalCategory,
              status: "Running",
              last_feed_company: trackForm.feed_company,
            };

            const { data: createdFarmer, error: createFarmerErr } =
              await supabase
                .from("farmers")
                .insert([newFarmerPayload])
                .select()
                .single();

            if (!createFarmerErr && createdFarmer) {
              targetFarmerId = createdFarmer.id;
            } else {
              delete newFarmerPayload.last_feed_company;
              const { data: retryFarmer } = await supabase
                .from("farmers")
                .insert([newFarmerPayload])
                .select()
                .single();
              if (retryFarmer) targetFarmerId = retryFarmer.id;
            }
          }
        }

        if (targetFarmerId) {
          const farmerUpdatePayload: Record<string, any> = {
            status: "Running",
            farm_type: finalCategory,
            last_feed_company: trackForm.feed_company,
          };
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

        // Insert into `doc_trackings`
        const insertPayload: Record<string, any> = {
          supply_id: activeTarget.supply_id || null,
          doc_id: activeTarget.doc_id,
          h_date: activeTarget.h_date,
          pcd: activeTarget.pcd,
          farmer_name: trackForm.farmer_name.trim(),
          address: trackForm.address.trim(),
          mobile: trackForm.mobile.trim(),
          qty: numQty,
          feed_company: trackForm.feed_company,
          remark: trackForm.remark.trim(),
          category: finalCategory,
          breed: finalBreed,
          farmer_id: targetFarmerId || null,
          created_by: currentUserId || null,
          tracked_by_name: currentUserName,
        };

        const { error: trkInsErr } = await supabase
          .from("doc_trackings")
          .insert([insertPayload]);

        if (trkInsErr) {
          delete insertPayload.category;
          delete insertPayload.breed;
          delete insertPayload.farmer_id;
          delete insertPayload.tracked_by_name;
          const { error: retryErr } = await supabase
            .from("doc_trackings")
            .insert([insertPayload]);
          if (retryErr) throw retryErr;
        }

        // Insert Running Batch into `vet_visits` (with (1), (2) indicator if already running)
        if (targetFarmerId && targetFeedAgentId) {
          const { data: runningBatches } = await supabase
            .from("vet_visits")
            .select("id, batch_label")
            .eq("farmer_id", String(targetFarmerId))
            .in("status", ["Activated", "Running"]);

          const runningCount = runningBatches?.length || 0;
          const baseFarmerName = trackForm.farmer_name.trim();
          const batchLabel =
            runningCount > 0
              ? `${baseFarmerName} (${runningCount + 1})`
              : baseFarmerName;

          if (
            runningCount === 1 &&
            runningBatches?.[0] &&
            !runningBatches[0].batch_label?.includes("(")
          ) {
            await supabase
              .from("vet_visits")
              .update({ batch_label: `${baseFarmerName} (1)` })
              .eq("id", runningBatches[0].id);
          }

          const nourishHatcheryName =
            companyNames.find((c) => c.toLowerCase().includes("nourish")) ||
            "Nourish";

          const visitPayload: Record<string, any> = {
            farmer_id: String(targetFarmerId),
            agent_id: String(targetFeedAgentId),
            company_id: String(targetCompId),
            category: finalCategory,
            hatch_date: activeTarget.h_date,
            hatchery: nourishHatcheryName,
            breed: finalBreed,
            box_no: "",
            t_birds: numQty,
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

      await loadAllData();
      setIsTrackModalOpen(false);
      setEditingTracking(null);
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
      await loadAllData();
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to delete tracking.");
    }
  }

  // ==================================================
  // FILTERED DATA & RESET
  // ==================================================
  const allPcdTypes = Array.from(new Set(supplies.map((s) => s.pcd))).sort();
  const allUpazilas = Array.from(new Set(agents.map((a) => a.upazila))).sort();

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

    setFromDate(firstDay);
    setToDate(lastDay);
    if (cachedPageState) {
      cachedPageState.monthFilter = ym;
      cachedPageState.fromDate = firstDay;
      cachedPageState.toDate = lastDay;
    }
  }

  function resetAllFilters() {
    const latestDate = supplies[0]?.h_date || "";
    setFromDate(latestDate);
    setToDate(latestDate);
    setMonthFilter("");
    setPcdFilter("");
    setUpazilaFilter("");
    setOfficerFilter("");
    setSearchAgent("");
    if (cachedPageState) {
      cachedPageState.fromDate = latestDate;
      cachedPageState.toDate = latestDate;
      cachedPageState.monthFilter = "";
    }
  }

  function matchesDateFilter(hDate: string) {
    if (monthFilter && !hDate.startsWith(monthFilter)) return false;
    if (fromDate && hDate < fromDate) return false;
    if (toDate && hDate > toDate) return false;
    return true;
  }

  const matchingAgents = agents.filter((agent) => {
    if (upazilaFilter && agent.upazila !== upazilaFilter) return false;

    if (officerFilter) {
      const selectedOff = officers.find((o) => o.id === officerFilter);
      if (selectedOff && !selectedOff.upazilas.includes(agent.upazila)) {
        return false;
      }
    }

    const q = searchAgent.trim().toLowerCase();
    if (q) {
      const matchSearch =
        agent.doc_id.toLowerCase().includes(q) ||
        agent.agent_name.toLowerCase().includes(q);
      if (!matchSearch) return false;
    }

    return true;
  });

  const matchingAgentDocIds = new Set(matchingAgents.map((a) => a.doc_id));

  const filteredSupplies = supplies.filter((s) => {
    if (!matchingAgentDocIds.has(s.doc_id)) return false;
    const matchPcd = !pcdFilter || s.pcd === pcdFilter;
    return matchesDateFilter(s.h_date) && matchPcd;
  });

  const filteredTrackings = trackings.filter((t) => {
    if (!matchingAgentDocIds.has(t.doc_id)) return false;
    const matchPcd = !pcdFilter || t.pcd === pcdFilter;
    return matchesDateFilter(t.h_date) && matchPcd;
  });

  const visibleAgents = matchingAgents.filter((agent) =>
    filteredSupplies.some((s) => s.doc_id === agent.doc_id)
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

  // ==================================================
  // MULTI-SELECT DOC TYPE EXCEL EXPORT
  // ==================================================
  function openExportModal() {
    setExportSelectedTypes(pcdFilter ? [pcdFilter] : [...allPcdTypes]);
    setIsExportModalOpen(true);
  }

  function toggleExportDocType(type: string) {
    setExportSelectedTypes((prev) =>
      prev.includes(type) ? prev.filter((t) => t !== type) : [...prev, type]
    );
  }

  function handleExportExcel() {
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
          s.doc_id === agent.doc_id && exportSelectedTypes.includes(s.pcd)
      );

      const groups = new Map<
        string,
        { h_date: string; pcd: string; qty: number }
      >();
      agentSupplies.forEach((s) => {
        const key = `${s.h_date}_${s.pcd}`;
        const prev = groups.get(key);
        if (prev) {
          prev.qty += Number(s.qty) || 0;
        } else {
          groups.set(key, {
            h_date: s.h_date,
            pcd: s.pcd,
            qty: Number(s.qty) || 0,
          });
        }
      });

      groups.forEach((g) => {
        grandTotalSupply += g.qty;

        const matchingTracks = filteredTrackings.filter(
          (t) =>
            t.doc_id === agent.doc_id &&
            t.h_date === g.h_date &&
            t.pcd === g.pcd
        );

        if (matchingTracks.length === 0) {
          exportRows.push({
            "H.Date": g.h_date,
            "DOC.ID": agent.doc_id,
            Agent: agent.agent_name,
            "Agent Address": `${agent.upazila}, ${agent.district}`,
            "DOC Qty": g.qty,
            "DOC Type": g.pcd,
            Farmer: "-",
            "Farmer Address": "-",
            "Farmer Mobile": "-",
            "Tracked Qty": 0,
            "Feed Company": "-",
            "Tracked By": "-",
            Remark: "Untracked",
          });
        } else {
          matchingTracks.forEach((t, index) => {
            grandTotalTracked += Number(t.qty) || 0;

            exportRows.push({
              "H.Date": g.h_date,
              "DOC.ID": agent.doc_id,
              Agent: agent.agent_name,
              "Agent Address": `${agent.upazila}, ${agent.district}`,
              "DOC Qty": index === 0 ? g.qty : "",
              "DOC Type": g.pcd,
              Farmer: t.farmer_name,
              "Farmer Address": t.address || "-",
              "Farmer Mobile": t.mobile || "-",
              "Tracked Qty": Number(t.qty) || 0,
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

    exportRows.push({
      "H.Date": "TOTAL",
      "DOC.ID": "",
      Agent: "",
      "Agent Address": "",
      "DOC Qty": grandTotalSupply,
      "DOC Type": "",
      Farmer: "",
      "Farmer Address": "",
      "Farmer Mobile": "",
      "Tracked Qty": grandTotalTracked,
      "Feed Company": "",
      "Tracked By": "",
      Remark: `Untracked: ${Math.max(0, grandTotalSupply - grandTotalTracked)}`,
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);

    const headers = Object.keys(exportRows[0]);
    worksheet["!cols"] = headers.map((header) => {
      let maxLen = header.length;
      exportRows.forEach((row) => {
        const val =
          row[header] !== undefined && row[header] !== null
            ? String(row[header])
            : "";
        if (val.length > maxLen) maxLen = val.length;
      });
      return { wch: Math.max(maxLen + 5, 12) };
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "DOC_Tracking");
    XLSX.writeFile(
      workbook,
      `DOC_Tracking_${dateRangeLabel}_${exportSelectedTypes.join("-")}.xlsx`
    );
    setIsExportModalOpen(false);
  }

  // ==================================================
  // TRACKED REPORT POPUP CALCULATIONS
  // ==================================================
  const reportSupplies = supplies.filter((s) => {
    if (!matchingAgentDocIds.has(s.doc_id)) return false;
    const matchFrom = !reportFromDate || s.h_date >= reportFromDate;
    const matchTo = !reportToDate || s.h_date <= reportToDate;
    const matchPcd = !reportPcdFilter || s.pcd === reportPcdFilter;
    return matchFrom && matchTo && matchPcd;
  });

  const reportTrackings = trackings.filter((t) => {
    if (!matchingAgentDocIds.has(t.doc_id)) return false;
    const matchFrom = !reportFromDate || t.h_date >= reportFromDate;
    const matchTo = !reportToDate || t.h_date <= reportToDate;
    const matchPcd = !reportPcdFilter || t.pcd === reportPcdFilter;
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
        if (!matchingAgentDocIds.has(s.doc_id)) return false;
        const matchFrom = !reportFromDate || s.h_date >= reportFromDate;
        const matchTo = !reportToDate || s.h_date <= reportToDate;
        return matchFrom && matchTo && s.pcd === pcd;
      })
      .reduce((sum, s) => sum + (Number(s.qty) || 0), 0);

    const trkQty = trackings
      .filter((t) => {
        if (!matchingAgentDocIds.has(t.doc_id)) return false;
        const matchFrom = !reportFromDate || t.h_date >= reportFromDate;
        const matchTo = !reportToDate || t.h_date <= reportToDate;
        return matchFrom && matchTo && t.pcd === pcd;
      })
      .reduce((sum, t) => sum + (Number(t.qty) || 0), 0);

    const untrkQty = Math.max(0, supQty - trkQty);
    const pct = supQty > 0 ? ((trkQty / supQty) * 100).toFixed(1) : "0.0";

    return { pcd, supQty, trkQty, untrkQty, pct };
  });

  const collaboratedFarmers = getCollaboratedFarmersForTarget(activeTarget);

  return (
    <div className="p-2.5 sm:p-4 max-w-7xl mx-auto space-y-3">
      {/* ==================================================
          COMPACT HEADER: STATS BADGES + ACTION BUTTONS
      ================================================== */}
      <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-2.5 bg-white border border-slate-200 rounded-xl px-3.5 py-2.5 shadow-sm">
        <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
          <span className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200 text-slate-700 text-xs font-medium">
            Date: <strong className="text-slate-900">{dateRangeLabel}</strong>
          </span>

          <span className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200 text-slate-700 text-xs font-medium">
            Total DOC:{" "}
            <strong className="text-slate-900">
              {totalSupplyQty.toLocaleString()}
            </strong>
          </span>

          <span className="px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-700 text-xs font-medium">
            Tracked:{" "}
            <strong>
              {totalTrackedQty.toLocaleString()} ({trackedPercent}%)
            </strong>
          </span>

          <span className="px-2.5 py-1 rounded-lg bg-amber-50 border border-amber-200 text-amber-700 text-xs font-medium">
            Untracked:{" "}
            <strong>
              {totalUntrackedQty.toLocaleString()} ({untrackedPercent}%)
            </strong>
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            onClick={() => initTrackingPage(true)}
            title="Refresh Data"
            className="p-2 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs transition"
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
                className="inline-flex items-center gap-1.5 bg-slate-700 hover:bg-slate-800 text-white px-3 py-1.5 rounded-lg text-xs font-medium transition"
              >
                {importing ? (
                  <FaSpinner className="animate-spin" />
                ) : (
                  <FaFileUpload />
                )}
                Import Excel
              </button>
            </>
          )}

          <button
            onClick={openExportModal}
            className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-medium transition"
          >
            <FaFileExcel />
            Export Excel
          </button>

          <button
            onClick={() => setIsReportOpen(true)}
            className="inline-flex items-center gap-1.5 bg-amber-500 hover:bg-amber-600 text-white px-3 py-1.5 rounded-lg text-xs font-medium transition"
          >
            <FaChartPie />
            Tracked Report
          </button>
        </div>
      </div>

      {/* ==================================================
          COMPACT FILTER BAR WITH SMALL RESET ICON
      ================================================== */}
      <div className="bg-white border border-slate-200 rounded-xl p-2.5 shadow-sm">
        <div
          className={`grid grid-cols-2 sm:grid-cols-4 ${
            isManagerOrAdmin ? "lg:grid-cols-7" : "lg:grid-cols-5"
          } gap-2 items-center`}
        >
          <div className="flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white focus-within:border-emerald-500">
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
              className="w-full text-xs outline-none bg-transparent text-slate-700"
            />
          </div>

          <div className="flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white focus-within:border-emerald-500">
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
              className="w-full text-xs outline-none bg-transparent text-slate-700"
            />
          </div>

          <select
            value={monthFilter}
            onChange={(e) => handleMonthChange(e.target.value)}
            className="border border-slate-300 rounded-lg px-2 py-1.5 text-xs bg-white outline-none focus:border-emerald-500"
          >
            <option value="">All Months</option>
            {allMonths.map((ym) => (
              <option key={ym} value={ym}>
                Month: {formatMonthLabel(ym)}
              </option>
            ))}
          </select>

          <select
            value={pcdFilter}
            onChange={(e) => setPcdFilter(e.target.value)}
            className="border border-slate-300 rounded-lg px-2 py-1.5 text-xs bg-white outline-none focus:border-emerald-500"
          >
            <option value="">All DOC Types</option>
            {allPcdTypes.map((pcd) => (
              <option key={pcd} value={pcd}>
                DOC Type: {pcd}
              </option>
            ))}
          </select>

          {isManagerOrAdmin && (
            <select
              value={upazilaFilter}
              onChange={(e) => setUpazilaFilter(e.target.value)}
              className="border border-slate-300 rounded-lg px-2 py-1.5 text-xs bg-white outline-none focus:border-emerald-500"
            >
              <option value="">All Upazilas</option>
              {allUpazilas.map((u) => (
                <option key={u} value={u}>
                  Upazila: {u}
                </option>
              ))}
            </select>
          )}

          {isManagerOrAdmin && (
            <select
              value={officerFilter}
              onChange={(e) => setOfficerFilter(e.target.value)}
              className="border border-slate-300 rounded-lg px-2 py-1.5 text-xs bg-white outline-none focus:border-emerald-500"
            >
              <option value="">All Officers</option>
              {officers
                .filter((o) => o.upazilas.length > 0)
                .map((off) => (
                  <option key={off.id} value={off.id}>
                    Officer: {off.full_name}
                  </option>
                ))}
            </select>
          )}

          <div className="col-span-2 sm:col-span-1 flex items-center gap-1.5">
            <input
              type="text"
              placeholder="Search ID or Agent..."
              value={searchAgent}
              onChange={(e) => setSearchAgent(e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs outline-none focus:border-emerald-500"
            />
            <button
              onClick={resetAllFilters}
              title="Reset Filters"
              className="shrink-0 inline-flex items-center justify-center rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2.5 py-1.5 text-xs font-semibold text-slate-700 transition"
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

      {/* AGENT-WISE CARDS */}
      {loading ? (
        <div className="bg-white border border-slate-200 rounded-2xl py-10 text-center text-slate-500 text-xs sm:text-sm">
          <FaSpinner className="animate-spin inline mr-2" />
          Loading DOC tracking data...
        </div>
      ) : visibleAgents.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-2xl py-12 text-center px-4">
          <p className="text-slate-600 text-sm sm:text-base font-bold">
            {fromDate || toDate || monthFilter
              ? "No DOC on this date"
              : "No DOC supply records found for the selected filter."}
          </p>
        </div>
      ) : (
        <div className="space-y-3.5">
          {visibleAgents.map((agent) => {
            const agentSupplies = filteredSupplies.filter(
              (s) => s.doc_id === agent.doc_id
            );

            const agentDocTypes = Array.from(
              new Set(agentSupplies.map((s) => s.pcd))
            );

            return (
              <div
                key={agent.id}
                className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden"
              >
                <div className="bg-slate-100 border-b border-slate-200 px-4 py-2.5 flex items-center gap-2.5">
                  <span className="px-2.5 py-0.5 rounded-md bg-slate-200/80 border border-slate-300 text-slate-800 font-bold text-xs">
                    {agent.doc_id}
                  </span>
                  <h2 className="font-bold text-slate-800 text-sm sm:text-base truncate">
                    {agent.agent_name}
                  </h2>
                </div>

                <div className="p-3 sm:p-4 space-y-4 divide-y divide-slate-100">
                  {agentDocTypes.map((docType) => {
                    const typeSupplyRows = agentSupplies.filter(
                      (s) => s.pcd === docType
                    );

                    const typeTotalQty = typeSupplyRows.reduce(
                      (sum, r) => sum + (Number(r.qty) || 0),
                      0
                    );

                    const typeTrackRows = filteredTrackings.filter(
                      (t) => t.doc_id === agent.doc_id && t.pcd === docType
                    );

                    const typeTrackedQty = typeTrackRows.reduce(
                      (sum, r) => sum + (Number(r.qty) || 0),
                      0
                    );

                    const typeUntrackedQty = Math.max(
                      0,
                      typeTotalQty - typeTrackedQty
                    );

                    const isAllTracked =
                      typeTotalQty > 0 && typeUntrackedQty <= 0;

                    const typeTrackedPct =
                      typeTotalQty > 0
                        ? ((typeTrackedQty / typeTotalQty) * 100).toFixed(1)
                        : "0.0";
                    const typeUntrackedPct =
                      typeTotalQty > 0
                        ? ((typeUntrackedQty / typeTotalQty) * 100).toFixed(1)
                        : "0.0";

                    const primaryDate =
                      typeSupplyRows[0]?.h_date || fromDate || todayStr;
                    const mappedInfo = mapPcdToCategoryAndBreed(docType);

                    return (
                      <div key={docType} className="pt-3 first:pt-0 space-y-2.5">
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex flex-wrap items-center gap-2">
                            <span className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200 text-slate-700 font-bold text-xs">
                              DOC Type: {docType}
                            </span>
                            <span className="px-2 py-0.5 rounded bg-emerald-50 border border-emerald-200 text-emerald-800 text-[11px] font-semibold">
                              {mappedInfo.category} • {mappedInfo.breed}
                            </span>
                          </div>

                          <button
                            type="button"
                            onClick={() => {
                              if (isAllTracked) {
                                alert("all DOC are tracked");
                                return;
                              }
                              openAddTrackingModal({
                                doc_id: agent.doc_id,
                                agent_name: agent.agent_name,
                                upazila: agent.upazila,
                                district: agent.district,
                                feed_agent_id: agent.feed_agent_id,
                                pcd: docType,
                                h_date: primaryDate,
                                supply_id: typeSupplyRows[0]?.id,
                                remainingQty: typeUntrackedQty,
                              });
                            }}
                            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition ${
                              isAllTracked
                                ? "bg-slate-200 text-slate-500 cursor-not-allowed"
                                : "bg-emerald-600 hover:bg-emerald-700 text-white"
                            }`}
                          >
                            {isAllTracked ? (
                              <span>✓ All DOC Tracked</span>
                            ) : (
                              <>
                                <FaPlus size={10} />
                                <span>Add Tracking</span>
                              </>
                            )}
                          </button>
                        </div>

                        {/* SUPPLY TABLE */}
                        <div className="overflow-x-auto border border-slate-200 rounded-xl">
                          <table className="min-w-full text-xs">
                            <thead className="bg-slate-50 border-b border-slate-200 text-slate-600">
                              <tr>
                                <th className="px-3 py-1.5 text-left font-semibold">
                                  H.Date
                                </th>
                                <th className="px-3 py-1.5 text-left font-semibold">
                                  DOC Type
                                </th>
                                <th className="px-3 py-1.5 text-left font-semibold">
                                  AG/B+
                                </th>
                                <th className="px-3 py-1.5 text-right font-semibold">
                                  Qty
                                </th>
                                <th className="px-3 py-1.5 text-right font-semibold">
                                  B.Qty
                                </th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {typeSupplyRows.map((row) => (
                                <tr key={row.id} className="hover:bg-slate-50">
                                  <td className="px-3 py-1.5 font-medium text-slate-800">
                                    {row.h_date}
                                  </td>
                                  <td className="px-3 py-1.5 font-bold text-slate-700">
                                    {row.pcd}
                                  </td>
                                  <td className="px-3 py-1.5 text-slate-600">
                                    {row.ag_b || "-"}
                                  </td>
                                  <td className="px-3 py-1.5 text-right font-bold text-slate-800">
                                    {Number(row.qty).toLocaleString()}
                                  </td>
                                  <td className="px-3 py-1.5 text-right text-slate-600">
                                    {Number(row.b_qty).toLocaleString()}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* TRACKED FARMERS LIST */}
                        {typeTrackRows.length > 0 && (
                          <div className="bg-emerald-50/40 border border-emerald-200 rounded-xl p-2 space-y-1.5">
                            <div className="overflow-x-auto">
                              <table className="min-w-full text-xs">
                                <thead className="border-b border-emerald-200 text-emerald-900">
                                  <tr>
                                    <th className="py-1.5 px-2 text-left">
                                      Farmer
                                    </th>
                                    <th className="py-1.5 px-2 text-left">
                                      Address
                                    </th>
                                    <th className="py-1.5 px-2 text-left">
                                      Mobile
                                    </th>
                                    <th className="py-1.5 px-2 text-left">
                                      Category & Breed
                                    </th>
                                    <th className="py-1.5 px-2 text-right">
                                      DOC Qty
                                    </th>
                                    <th className="py-1.5 px-2 text-left">
                                      Feed Company
                                    </th>
                                    <th className="py-1.5 px-2 text-left">
                                      Tracked By
                                    </th>
                                    <th className="py-1.5 px-2 text-left">
                                      Remark
                                    </th>
                                    <th className="py-1.5 px-2 text-right">
                                      Action
                                    </th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-emerald-100">
                                  {typeTrackRows.map((tr) => {
                                    const allowedToEdit = canModifyTracking(tr);
                                    const autoInfo = mapPcdToCategoryAndBreed(
                                      tr.pcd
                                    );
                                    return (
                                      <tr key={tr.id}>
                                        <td className="py-1.5 px-2 font-semibold text-slate-800">
                                          {tr.farmer_name}
                                        </td>
                                        <td className="py-1.5 px-2 text-slate-600">
                                          {tr.address || "-"}
                                        </td>
                                        <td className="py-1.5 px-2 text-slate-600">
                                          {tr.mobile || "-"}
                                        </td>
                                        <td className="py-1.5 px-2 text-slate-700 font-medium">
                                          {tr.category || autoInfo.category} •{" "}
                                          {tr.breed || autoInfo.breed}
                                        </td>
                                        <td className="py-1.5 px-2 text-right font-bold text-emerald-700">
                                          {Number(tr.qty).toLocaleString()}
                                        </td>
                                        <td className="py-1.5 px-2">
                                          <span className="px-2 py-0.5 rounded bg-white border border-emerald-300 text-emerald-800 font-semibold">
                                            {tr.feed_company}
                                          </span>
                                        </td>
                                        <td className="py-1.5 px-2 font-medium text-slate-700">
                                          {tr.tracked_by_name || "-"}
                                        </td>
                                        <td className="py-1.5 px-2 text-slate-600">
                                          {tr.remark || "-"}
                                        </td>
                                        <td className="py-1.5 px-2 text-right whitespace-nowrap space-x-1">
                                          {allowedToEdit ? (
                                            <>
                                              <button
                                                type="button"
                                                onClick={() =>
                                                  openEditTrackingModal(
                                                    tr,
                                                    agent,
                                                    typeUntrackedQty
                                                  )
                                                }
                                                className="text-blue-600 hover:text-blue-800 p-1"
                                                title="Edit Tracking"
                                              >
                                                <FaEdit size={12} />
                                              </button>
                                              <button
                                                type="button"
                                                onClick={() =>
                                                  handleDeleteTracking(tr)
                                                }
                                                className="text-red-500 hover:text-red-700 p-1"
                                                title="Delete Tracking"
                                              >
                                                <FaTrash size={11} />
                                              </button>
                                            </>
                                          ) : (
                                            <span className="text-[10px] text-slate-400 italic">
                                              View only
                                            </span>
                                          )}
                                        </td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}

                        <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs">
                          <span className="font-medium text-slate-700">
                            <strong>{docType}</strong> Total:{" "}
                            <strong>{typeTotalQty.toLocaleString()}</strong>
                          </span>

                          <div className="flex flex-wrap items-center gap-3">
                            <span className="text-emerald-700 font-semibold">
                              Tracked: {typeTrackedQty.toLocaleString()} (
                              {typeTrackedPct}%)
                            </span>
                            <span className="text-amber-700 font-semibold">
                              Untracked: {typeUntrackedQty.toLocaleString()} (
                              {typeUntrackedPct}%)
                            </span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* ==================================================
          MODAL 1: ADD / EDIT FARMER TRACKING
      ================================================== */}
      {isTrackModalOpen && activeTarget && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-white rounded-t-2xl sm:rounded-2xl shadow-2xl w-full max-w-md max-h-[92vh] flex flex-col overflow-hidden">
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-slate-200 bg-slate-100">
              <div>
                <h3 className="text-sm sm:text-base font-bold text-slate-800">
                  {editingTracking ? "Edit Tracking" : "Add Tracking"} — DOC
                  Type: {activeTarget.pcd}
                </h3>
                <p className="text-xs text-slate-600">
                  {activeTarget.doc_id} • {activeTarget.agent_name} (Remaining:{" "}
                  <strong>{activeTarget.remainingQty}</strong>)
                </p>
              </div>
              <button
                onClick={() => {
                  setIsTrackModalOpen(false);
                  setEditingTracking(null);
                }}
                className="w-8 h-8 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes />
              </button>
            </div>

            <form
              onSubmit={handleSaveTracking}
              className="p-5 space-y-3.5 overflow-y-auto"
            >
              {/* Select Farmer from List Dropdown right above Farmer Name */}
              <div>
                <label className="block text-xs font-semibold text-emerald-700 mb-1">
                  Select Farmer from List (Optional)
                </label>
                <select
                  value={selectedFarmerId}
                  onChange={(e) =>
                    handleSelectFarmerFromDropdown(e.target.value)
                  }
                  className="w-full border border-emerald-300 rounded-xl px-3 py-2 text-xs sm:text-sm bg-emerald-50/40 text-slate-800 outline-none focus:border-emerald-600"
                >
                  <option value="">
                    -- Select from List ({collaboratedFarmers.length} farmers)
                    or type below --
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

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Farmer Name *
                </label>
                <input
                  type="text"
                  required
                  placeholder="Enter farmer name"
                  value={trackForm.farmer_name}
                  onChange={(e) =>
                    setTrackForm({ ...trackForm, farmer_name: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Address
                  </label>
                  <input
                    type="text"
                    placeholder="Village / Area"
                    value={trackForm.address}
                    onChange={(e) =>
                      setTrackForm({ ...trackForm, address: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Mobile
                  </label>
                  <input
                    type="text"
                    placeholder="017XXXXXXXX"
                    value={trackForm.mobile}
                    onChange={(e) =>
                      setTrackForm({ ...trackForm, mobile: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              {/* Farm Category & Breed Auto-filled from PCD */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
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
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-xs sm:text-sm bg-white outline-none focus:border-emerald-500"
                  >
                    {FARM_CATEGORIES.map((cat) => (
                      <option key={cat} value={cat}>
                        {cat}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Breed ({activeTarget.pcd}) *
                  </label>
                  <select
                    value={trackForm.breed}
                    onChange={(e) =>
                      setTrackForm({ ...trackForm, breed: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-xs sm:text-sm bg-white outline-none focus:border-emerald-500"
                  >
                    {(BREEDS_BY_CATEGORY[trackForm.category] || []).map((b) => (
                      <option key={b} value={b}>
                        {b}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    DOC Qty *
                  </label>
                  <input
                    type="number"
                    required
                    min={1}
                    max={activeTarget.remainingQty}
                    placeholder="Qty"
                    value={trackForm.qty}
                    onChange={(e) =>
                      setTrackForm({ ...trackForm, qty: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
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
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none bg-white focus:border-emerald-500"
                  >
                    <option value="">Select Company</option>
                    {companyNames.map((fc) => (
                      <option key={fc} value={fc}>
                        {fc}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Remark
                </label>
                <input
                  type="text"
                  placeholder="Optional note..."
                  value={trackForm.remark}
                  onChange={(e) =>
                    setTrackForm({ ...trackForm, remark: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                />
              </div>

              <div className="flex gap-2.5 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setIsTrackModalOpen(false);
                    setEditingTracking(null);
                  }}
                  className="flex-1 py-2 rounded-xl border border-slate-300 text-slate-600 text-sm font-medium"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={submitting}
                  className="flex-1 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium flex items-center justify-center gap-2"
                >
                  {submitting && <FaSpinner className="animate-spin" />}
                  {editingTracking ? "Update Tracking" : "Save Tracking"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 2: EXPORT EXCEL (Multi-Select DOC Type)
      ================================================== */}
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
                      exportSelectedTypes.length === allPcdTypes.length
                        ? []
                        : [...allPcdTypes]
                    )
                  }
                  className="text-emerald-600 font-semibold hover:underline"
                >
                  {exportSelectedTypes.length === allPcdTypes.length
                    ? "Deselect All"
                    : "Select All"}
                </button>
              </div>

              <div className="border border-slate-200 rounded-xl max-h-48 overflow-y-auto p-2 space-y-1 bg-slate-50">
                {allPcdTypes.map((type) => {
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
                      <span>DOC Type: {type}</span>
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

      {/* ==================================================
          MODAL 3: TRACKED REPORT POPUP
      ================================================== */}
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
                        {p}
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
                            {p.pcd}
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