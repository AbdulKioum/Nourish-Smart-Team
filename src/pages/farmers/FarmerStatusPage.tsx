import { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

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
  Broiler: ["Arber Acress", "Ross", "Cobb 500", "Indian River", "Hubbard", "Hubbard EP"],
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

function normalizeCategory(f: any): string {
  const raw = String(f.farm_type || f.farm_category || f.category || "").trim();
  const matched = CATEGORIES.find(
    (c) => c.toLowerCase() === raw.toLowerCase()
  );
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
  const diff = new Date().getTime() - d.getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function calcAgeDays(hatchDate?: string, visitDate?: string): number {
  if (!hatchDate) return 0;
  const hd = new Date(hatchDate);
  const vd = visitDate ? new Date(visitDate) : new Date();
  return Math.max(0, Math.floor((vd.getTime() - hd.getTime()) / 86400000));
}

export default function FarmerStatusPage() {
  const [companies, setCompanies] = useState<any[]>([]);
  const [selectedCompany, setSelectedCompany] = useState<string>("all");
  const [agents, setAgents] = useState<any[]>([]);
  const [selectedAgent, setSelectedAgent] = useState<string>("all");
  const [selectedCategory, setSelectedCategory] = useState<string>("all");
  const [selectedStatusFilter, setSelectedStatusFilter] =
    useState<StatusTab>("all");
  const [searchQuery, setSearchQuery] = useState<string>("");

  const [expandedCard, setExpandedCard] = useState<string | null>(null);
  const [enrichedFarmers, setEnrichedFarmers] = useState<any[]>([]);
  const [allVisitsMap, setAllVisitsMap] = useState<Record<string, any>>({});
  const [broilerStds, setBroilerStds] = useState<Record<number, any>>({});
  const [layerStds, setLayerStds] = useState<Record<number, any>>({});
  const [loading, setLoading] = useState(true);

  const [selectedFarmerIds, setSelectedFarmerIds] = useState<string[]>([]);
  const [isBulkMode, setIsBulkMode] = useState(false);

  const [showModal, setShowModal] = useState(false);
  const [selectedFarmer, setSelectedFarmer] = useState<any>(null);
  const [form, setForm] = useState({
    category: "Broiler",
    hatch_date: new Date().toISOString().split("T")[0],
    hatchery: "",
    breed: "",
    box_no: "",
    t_birds: "",
  });

  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [historyFarmer, setHistoryFarmer] = useState<any>(null);
  const [historyData, setHistoryData] = useState<{
    before: any[];
    after: any[];
    layer: any[];
    feedLogs: any[];
  }>({ before: [], after: [], layer: [], feedLogs: [] });

  useEffect(() => {
    loadAllStatusData();
  }, []);

  async function loadAllStatusData() {
    setLoading(true);
    try {
      const [
        { data: compData },
        { data: agData },
        { data: frmData },
        { data: visitData },
        { data: bStdData },
        { data: lStdData },
      ] = await Promise.all([
        supabase.from("companies").select("*").order("company_name"),
        supabase.from("agents").select("*").order("agent_name"),
        supabase.from("farmers").select("*").order("farmer_name"),
        supabase
          .from("vet_visits")
          .select("*")
          .order("created_at", { ascending: false }),
        supabase.from("broiler_standards").select("*"),
        supabase.from("layer_standards").select("*"),
      ]);

      const bMap: Record<number, any> = {};
      (bStdData || []).forEach((s: any) => (bMap[Number(s.age_days)] = s));
      setBroilerStds(bMap);

      const lMap: Record<number, any> = {};
      (lStdData || []).forEach((s: any) => (lMap[Number(s.age_weeks)] = s));
      setLayerStds(lMap);

      const vMap: Record<string, any> = {};
      (visitData || []).forEach((v: any) => (vMap[String(v.id)] = v));
      setAllVisitsMap(vMap);

      const allCompanies = compData || [];
      const allAgents = agData || [];
      const allFarmers = frmData || [];
      const allVisits = visitData || [];

      setCompanies(allCompanies);
      setAgents(allAgents);

      if (selectedCompany === "all" && allCompanies.length > 0) {
        const nourish = allCompanies.find((c) =>
          c.company_name?.toLowerCase().includes("nourish")
        );
        if (nourish) setSelectedCompany(String(nourish.id));
      }

      const processed = allFarmers.map((f) => {
        const farmerVisits = allVisits.filter(
          (v) => String(v.farmer_id) === String(f.id)
        );
        const runningVisits = farmerVisits.filter(
          (v) => v.status === "Activated" || v.status === "Running"
        );
        const closedVisits = farmerVisits.filter(
          (v) => v.status === "Sold" || v.status === "Culled"
        );

        const latestRunning = runningVisits[0] || null;
        const latestClosed = closedVisits[0] || null;

        const lastSaleDate =
          f.last_sale_date ||
          latestClosed?.closed_date ||
          (latestClosed?.created_at
            ? String(latestClosed.created_at).slice(0, 10)
            : null);

        const daysSinceSale = getDaysSince(lastSaleDate);

        let computedStatus: "Running" | "Active" | "Inactive" = "Inactive";
        let statusDate = "";

        if (runningVisits.length > 0) {
          computedStatus = "Running";
          statusDate = latestRunning?.hatch_date || "";
        } else if (daysSinceSale !== null && daysSinceSale <= 60) {
          computedStatus = "Active";
          statusDate = lastSaleDate || "";
        } else {
          computedStatus = "Inactive";
          statusDate = lastSaleDate || "";
        }

        const cat = latestRunning?.category || normalizeCategory(f);
        const agentObj = allAgents.find(
          (a) => String(a.id) === String(f.agent_id)
        );
        const compObj = allCompanies.find(
          (c) => String(c.id) === String(f.company_id)
        );

        return {
          ...f,
          computedStatus,
          statusDate,
          category: cat,
          runningCount: runningVisits.length,
          runningVisits,
          latestRunning,
          latestVisit: latestRunning || latestClosed || null,
          agentName: agentObj?.agent_name || "Unassigned Agent",
          companyName: compObj?.company_name || "-",
          nourishTag: f.nourish_status || "",
          nourishTagDate: f.nourish_status_date || "",
        };
      });

      setEnrichedFarmers(processed);
    } catch (err) {
      console.error("Failed to load farmer status:", err);
    } finally {
      setLoading(false);
    }
  }

  const companyScopedFarmers = enrichedFarmers.filter((f) => {
    const matchComp =
      selectedCompany === "all" ||
      String(f.company_id) === String(selectedCompany);
    const matchAgent =
      selectedAgent === "all" || String(f.agent_id) === String(selectedAgent);
    return matchComp && matchAgent;
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

  const summaryCards = [
    {
      key: "all" as StatusTab,
      title: "Total Farmers",
      color: "border-slate-300 bg-white text-slate-800",
      badgeBg: "bg-slate-800 text-white",
      data: getSummaryForGroup(() => true),
    },
    {
      key: "Running" as StatusTab,
      title: "Running",
      color: "border-emerald-300 bg-emerald-50/60 text-emerald-900",
      badgeBg: "bg-emerald-600 text-white",
      data: getSummaryForGroup((f) => f.computedStatus === "Running"),
    },
    {
      key: "Active" as StatusTab,
      title: "Active (< 2M)",
      color: "border-blue-300 bg-blue-50/60 text-blue-900",
      badgeBg: "bg-blue-600 text-white",
      data: getSummaryForGroup((f) => f.computedStatus === "Active"),
    },
    {
      key: "Inactive" as StatusTab,
      title: "Inactive (> 2M)",
      color: "border-amber-300 bg-amber-50/60 text-amber-900",
      badgeBg: "bg-amber-600 text-white",
      data: getSummaryForGroup((f) => f.computedStatus === "Inactive"),
    },
    {
      key: "Leave from Nourish" as StatusTab,
      title: "Leave from Nourish",
      color: "border-red-300 bg-red-50/60 text-red-900",
      badgeBg: "bg-red-600 text-white",
      data: getSummaryForGroup((f) => f.nourishTag === "Leave from Nourish"),
    },
    {
      key: "Convert to Nourish" as StatusTab,
      title: "Convert to Nourish",
      color: "border-teal-300 bg-teal-50/60 text-teal-900",
      badgeBg: "bg-teal-600 text-white",
      data: getSummaryForGroup((f) => f.nourishTag === "Convert to Nourish"),
    },
  ];

  const tableFilteredFarmers = companyScopedFarmers.filter((f) => {
    if (
      selectedCategory !== "all" &&
      f.category?.toLowerCase() !== selectedCategory.toLowerCase()
    ) {
      return false;
    }

    if (selectedStatusFilter !== "all") {
      if (
        selectedStatusFilter === "Leave from Nourish" ||
        selectedStatusFilter === "Convert to Nourish"
      ) {
        if (f.nourishTag !== selectedStatusFilter) return false;
      } else {
        if (f.computedStatus !== selectedStatusFilter) return false;
      }
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      const match =
        f.farmer_name?.toLowerCase().includes(q) ||
        f.mobile?.toLowerCase().includes(q) ||
        f.village?.toLowerCase().includes(q) ||
        f.agentName?.toLowerCase().includes(q);
      if (!match) return false;
    }

    return true;
  });

  const groupedByAgent = Object.values(
    tableFilteredFarmers.reduce((acc: Record<string, any>, farmer: any) => {
      const key = String(farmer.agent_id || "unknown");
      if (!acc[key]) {
        acc[key] = {
          agentId: key,
          agentName: farmer.agentName,
          companyName: farmer.companyName,
          farmers: [],
        };
      }
      acc[key].farmers.push(farmer);
      return acc;
    }, {})
  );

  // Excel Export for Filtered Farmer Status List
  async function exportFilteredFarmersExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Farmer Status");

    sheet.addRow([
      "SL",
      "Company",
      "Agent",
      "Farmer Name",
      "Mobile",
      "Address",
      "Category",
      "Status",
      "Date",
      "Nourish Conversion",
    ]);

    sheet.getRow(1).eachCell((c) => {
      c.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "0F172A" },
      };
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    });

    tableFilteredFarmers.forEach((f, idx) => {
      sheet.addRow([
        idx + 1,
        f.companyName,
        f.agentName,
        f.farmer_name,
        f.mobile || "-",
        [f.village, f.upazila].filter(Boolean).join(", "),
        f.category,
        f.computedStatus,
        f.statusDate || "-",
        f.nourishTag
          ? `${f.nourishTag}${
              f.nourishTagDate ? ` (${f.nourishTagDate})` : ""
            }`
          : "-",
      ]);
    });

    sheet.columns.forEach((col) => (col.width = 20));
    const buf = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buf]),
      `Farmer_Status_${selectedStatusFilter}_${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`
    );
  }

  // Excel Export for Single Farmer History Modal
  async function exportSingleFarmerHistoryExcel() {
    if (!historyFarmer) return;
    const workbook = new ExcelJS.Workbook();
    const isLayer = (historyFarmer.category || "").toLowerCase() === "layer";

    if (isLayer) {
      const sheet = workbook.addWorksheet("Layer History");
      sheet.addRow([
        "Visit Date",
        "H.Date",
        "T.Birds",
        "Age (Wks)",
        "Weight (g)",
        "Std Wt (g)",
        "Feed (g/b/d)",
        "Std Feed (g)",
        "Egg Prod (%)",
        "Std Egg (%)",
        "Egg Wt (g)",
        "Std Egg Wt",
        "Feed Company",
        "Feed Type",
        "Remarks",
      ]);
      historyData.layer.forEach((r: any) => {
        const v = allVisitsMap[String(r.vet_visit_id)] || historyFarmer.latestVisit;
        const hDate = v?.hatch_date || "";
        const tBirds = v?.t_birds || 0;
        const ageWks = Math.max(1, Math.floor(calcAgeDays(hDate, r.visit_date) / 7));
        const std = layerStds[ageWks];
        sheet.addRow([
          r.visit_date || "-",
          hDate || "-",
          tBirds,
          ageWks,
          r.b_weight_g,
          std?.weight || "-",
          r.feed_intake_g,
          std?.feed_bird_day || "-",
          r.egg_production,
          std?.egg_production || "-",
          r.egg_weight,
          std?.egg_weight || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          r.remarks || "-",
        ]);
      });
    } else {
      const bSheet = workbook.addWorksheet("Before Sale History");
      bSheet.addRow([
        "Visit Date",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt (g)",
        "Feed (Kg)",
        "Feed (g/b)",
        "Std Feed (g)",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Company",
        "Feed Type",
        "Remarks",
      ]);
      historyData.before.forEach((r: any) => {
        const v = allVisitsMap[String(r.vet_visit_id)] || historyFarmer.latestVisit;
        const hDate = v?.hatch_date || "";
        const tBirds = v?.t_birds || 0;
        const ageDays = calcAgeDays(hDate, r.visit_date);
        const std = broilerStds[ageDays];
        bSheet.addRow([
          r.visit_date || "-",
          hDate || "-",
          tBirds,
          ageDays,
          r.mortality_pcs,
          r.weight_gm_b,
          std?.weight || "-",
          r.feed_t_kg,
          r.feed_gm_b,
          std?.feed_intake || "-",
          r.fcr_w_out_m,
          std?.fcr || "-",
          r.fcr_with_m,
          r.feed_company || "-",
          r.feed_type || "-",
          r.remarks || "-",
        ]);
      });

      const aSheet = workbook.addWorksheet("After Sale History");
      aSheet.addRow([
        "Sale Date",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt (g)",
        "Total Feed (Kg)",
        "Feed/Bird (g)",
        "Std Feed (g)",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Mill",
        "Feed Type",
        "Remarks",
      ]);
      historyData.after.forEach((r: any) => {
        const v = allVisitsMap[String(r.vet_visit_id)] || historyFarmer.latestVisit;
        const hDate = v?.hatch_date || "";
        const tBirds = Number(v?.t_birds || 0);
        const ageDays = calcAgeDays(hDate, r.visit_date);
        const std = broilerStds[ageDays];
        const feedGmB =
          r.feed_gm_b ||
          (tBirds > 0
            ? ((Number(r.feed_consumption_kg || 0) * 1000) / tBirds).toFixed(1)
            : "-");
        aSheet.addRow([
          r.visit_date || "-",
          hDate || "-",
          tBirds,
          ageDays,
          r.mortality,
          r.weight,
          std?.weight || "-",
          r.feed_consumption_kg,
          feedGmB,
          std?.feed_intake || "-",
          r.fcr_w_out_m,
          std?.fcr || "-",
          r.fcr_with_m,
          r.feed_mill || "-",
          r.feed || "-",
          r.remarks || "-",
        ]);
      });
    }

    const buf = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buf]),
      `History_${historyFarmer.farmer_name}_${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`
    );
  }

  function toggleFarmerCheckbox(id: string) {
    setSelectedFarmerIds((prev) =>
      prev.includes(id) ? prev.filter((item) => item !== id) : [...prev, id]
    );
  }

  function handleOpenSingleActivate(farmer: any) {
    const defaultCat = farmer.category || "Broiler";
    const defaultBreeds = BREEDS_BY_CATEGORY[defaultCat] || [];
    const defaultHatchery =
      companies.find((c) => String(c.id) === String(farmer.company_id))
        ?.company_name ||
      companies[0]?.company_name ||
      "";

    setIsBulkMode(false);
    setSelectedFarmer(farmer);
    setForm({
      category: defaultCat,
      hatch_date: new Date().toISOString().split("T")[0],
      hatchery: defaultHatchery,
      breed: defaultBreeds[0] || "",
      box_no: "",
      t_birds: "",
    });
    setShowModal(true);
  }

  function handleOpenBulkActivate() {
    if (selectedFarmerIds.length === 0) return;
    const defaultHatchery =
      companies.find((c) => String(c.id) === String(selectedCompany))
        ?.company_name ||
      companies[0]?.company_name ||
      "";

    setIsBulkMode(true);
    setSelectedFarmer(null);
    setForm({
      category: "Broiler",
      hatch_date: new Date().toISOString().split("T")[0],
      hatchery: defaultHatchery,
      breed: BREEDS_BY_CATEGORY["Broiler"][0],
      box_no: "",
      t_birds: "",
    });
    setShowModal(true);
  }

  async function handleActivateSubmit() {
    if (!form.hatch_date || !form.t_birds) {
      alert("Hatch Date and Total Birds are required.");
      return;
    }

    const targets = isBulkMode
      ? enrichedFarmers.filter((f) => selectedFarmerIds.includes(String(f.id)))
      : selectedFarmer
      ? [selectedFarmer]
      : [];

    if (targets.length === 0) return;

    try {
      for (const target of targets) {
        const nextBatchNum = (target.runningCount || 0) + 1;
        const baseName = target.farmer_name || "Farmer";
        const batchLabel =
          nextBatchNum > 1 ? `${baseName} (${nextBatchNum})` : baseName;

        const selectedCompObj = companies.find(
          (c) => String(c.id) === String(target.company_id)
        );
        const isNourishComp = Boolean(
          selectedCompObj?.company_name?.toLowerCase().includes("nourish")
        );

        const farmerUpdatePayload: any = {
          status: "Running",
          farm_type: form.category,
        };
        if (isNourishComp && !target.last_feed_company) {
          farmerUpdatePayload.last_feed_company = "Nourish";
        }

        await supabase
          .from("farmers")
          .update(farmerUpdatePayload)
          .eq("id", target.id);

        if (
          nextBatchNum === 2 &&
          target.runningVisits?.[0] &&
          !target.runningVisits[0].batch_label?.includes("(")
        ) {
          await supabase
            .from("vet_visits")
            .update({ batch_label: `${baseName} (1)` })
            .eq("id", target.runningVisits[0].id);
        }

        const insertPayload: any = {
          farmer_id: String(target.id),
          agent_id: String(target.agent_id),
          company_id: String(target.company_id),
          category: form.category,
          hatch_date: form.hatch_date,
          hatchery: form.hatchery,
          breed: form.breed,
          box_no: form.box_no,
          t_birds: Number(form.t_birds),
          status: "Activated",
          batch_label: batchLabel,
        };

        const { error: visitErr } = await supabase
          .from("vet_visits")
          .insert([insertPayload]);

        if (visitErr) {
          delete insertPayload.batch_label;
          const { error: retryErr } = await supabase
            .from("vet_visits")
            .insert([insertPayload]);
          if (retryErr) throw retryErr;
        }
      }

      alert("Activated & moved to Running Vet Visits!");
      setShowModal(false);
      setSelectedFarmerIds([]);
      loadAllStatusData();
    } catch (err: any) {
      alert(err.message || "Failed to activate farmer");
    }
  }

  async function handleOpenHistory(farmer: any) {
    setHistoryFarmer(farmer);
    setShowHistoryModal(true);

    const [{ data: bData }, { data: aData }, { data: lData }, { data: fLogs }] =
      await Promise.all([
        supabase
          .from("before_sale_reports")
          .select("*")
          .eq("farmer_id", String(farmer.id))
          .order("visit_date", { ascending: false }),
        supabase
          .from("after_sale_reports")
          .select("*")
          .eq("farmer_id", String(farmer.id))
          .order("visit_date", { ascending: false }),
        supabase
          .from("layer_reports")
          .select("*")
          .eq("farmer_id", String(farmer.id))
          .order("visit_date", { ascending: false }),
        supabase
          .from("farmer_feed_logs")
          .select("*")
          .eq("farmer_id", String(farmer.id))
          .order("change_date", { ascending: false }),
      ]);

    setHistoryData({
      before: bData || [],
      after: aData || [],
      layer: lData || [],
      feedLogs: fLogs || [],
    });
  }

  const availableAgentsForDropdown =
    selectedCompany === "all"
      ? agents
      : agents.filter((a) => String(a.company_id) === String(selectedCompany));

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-4">
      {/* Header + Export & Bulk Activate */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-4 py-2.5 shadow-sm">
        <div className="flex items-center gap-3">
          <h1 className="text-base sm:text-lg font-bold text-slate-800">
            Farmer Status
          </h1>
          <span className="text-xs font-semibold bg-slate-100 border border-slate-200 px-2.5 py-0.5 rounded-lg text-slate-700">
            Total: <strong>{tableFilteredFarmers.length}</strong>
          </span>
        </div>

        <div className="flex items-center gap-2">
          {selectedFarmerIds.length > 0 && (
            <button
              type="button"
              onClick={handleOpenBulkActivate}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow-sm transition"
            >
              Activate Selected ({selectedFarmerIds.length})
            </button>
          )}
          <button
            type="button"
            onClick={exportFilteredFarmersExcel}
            className="bg-slate-800 hover:bg-slate-900 text-white px-3 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            Export Excel
          </button>
        </div>
      </div>

      {/* ==================================================
          1. ULTRA-MINIMAL SUMMARY CARDS (NO EXTRA ICONS OR "FILTER" TEXT)
      ================================================== */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2.5">
        {summaryCards.map((card) => {
          const isSelected = selectedStatusFilter === card.key;
          const isExpanded = expandedCard === card.key;

          return (
            <div
              key={card.key}
              onClick={() => setSelectedStatusFilter(card.key)}
              className={`rounded-xl border p-3 shadow-sm cursor-pointer transition flex flex-col justify-between ${
                card.color
              } ${isSelected ? "ring-2 ring-emerald-600" : ""}`}
            >
              <div className="flex items-center justify-between gap-1">
                <span className="text-xs font-bold leading-tight">
                  {card.title}
                </span>
                <span
                  className={`px-2 py-0.5 rounded-full text-xs font-extrabold ${card.badgeBg}`}
                >
                  {card.data.total}
                </span>
              </div>

              <div className="mt-2 pt-1 border-t border-black/10 flex justify-end">
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setExpandedCard(isExpanded ? null : card.key);
                  }}
                  className="text-[11px] font-semibold hover:underline opacity-85"
                >
                  {isExpanded ? "Hide Details" : "Details"}
                </button>
              </div>

              {isExpanded && (
                <div
                  onClick={(e) => e.stopPropagation()}
                  className="mt-2 pt-2 border-t border-black/10 space-y-1 text-[11px] bg-white/95 rounded-lg p-2 text-slate-800"
                >
                  {CATEGORIES.map((cat) => (
                    <div
                      key={cat}
                      onClick={() => {
                        setSelectedStatusFilter(card.key);
                        setSelectedCategory(cat);
                      }}
                      className="flex justify-between items-center cursor-pointer hover:bg-slate-100 px-1 rounded"
                    >
                      <span>{cat}:</span>
                      <span className="font-bold">
                        {card.data.byCat[cat] || 0}
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* ==================================================
          2. COMPACT FILTER BAR BELOW SUMMARY
      ================================================== */}
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 items-center">
          <input
            type="text"
            placeholder="Search farmer, mobile..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="col-span-2 sm:col-span-1 w-full rounded-lg border border-slate-300 px-2.5 py-1.5 text-xs outline-none focus:border-emerald-500"
          />

          <select
            value={selectedCompany}
            onChange={(e) => {
              setSelectedCompany(e.target.value);
              setSelectedAgent("all");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs outline-none"
          >
            <option value="all">All Companies</option>
            {companies.map((comp) => (
              <option key={comp.id} value={String(comp.id)}>
                {comp.company_name}
              </option>
            ))}
          </select>

          <select
            value={selectedAgent}
            onChange={(e) => setSelectedAgent(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs outline-none"
          >
            <option value="all">All Agents</option>
            {availableAgentsForDropdown.map((ag) => (
              <option key={ag.id} value={String(ag.id)}>
                {ag.agent_name}
              </option>
            ))}
          </select>

          <select
            value={selectedCategory}
            onChange={(e) => setSelectedCategory(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs outline-none"
          >
            <option value="all">All Categories</option>
            {CATEGORIES.map((cat) => (
              <option key={cat} value={cat}>
                {cat}
              </option>
            ))}
          </select>

          <select
            value={selectedStatusFilter}
            onChange={(e) =>
              setSelectedStatusFilter(e.target.value as StatusTab)
            }
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs outline-none font-semibold"
          >
            <option value="all">All Status</option>
            <option value="Running">Running</option>
            <option value="Active">Active (&lt; 2 Months)</option>
            <option value="Inactive">Inactive (&gt; 2 Months)</option>
            <option value="Leave from Nourish">Leave from Nourish</option>
            <option value="Convert to Nourish">Convert to Nourish</option>
          </select>

          <button
            type="button"
            onClick={() => {
              setSelectedCompany("all");
              setSelectedAgent("all");
              setSelectedCategory("all");
              setSelectedStatusFilter("all");
              setSearchQuery("");
              setSelectedFarmerIds([]);
            }}
            className="rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 transition"
          >
            Reset
          </button>
        </div>
      </div>

      {/* ==================================================
          3. AGENT-WISE MINIMAL BORDERED FARMER TABLE (NOURISH CONVERSION MERGED INTO STATUS)
      ================================================== */}
      {loading ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-500">
          Loading Farmer Status...
        </div>
      ) : groupedByAgent.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-400">
          No farmers found for the selected filter.
        </div>
      ) : (
        <div className="space-y-4">
          {groupedByAgent.map((group: any) => (
            <div
              key={group.agentId}
              className="bg-white border border-slate-300 rounded-xl overflow-hidden shadow-sm"
            >
              <div className="bg-slate-100 border-b border-slate-300 px-4 py-2 flex flex-wrap items-center justify-between gap-2">
                <div className="text-xs sm:text-sm font-bold text-slate-800">
                  Agent:{" "}
                  <span className="text-emerald-700">{group.agentName}</span>{" "}
                  <span className="font-normal text-slate-500">
                    ({group.companyName})
                  </span>
                </div>
                <span className="text-xs font-semibold bg-white border border-slate-300 px-2.5 py-0.5 rounded text-slate-700">
                  Farmers: {group.farmers.length}
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-xs">
                  <thead className="bg-slate-50">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1.5 text-center w-8">
                        ✓
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                        SL
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                        Farmer Name
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                        Mobile
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                        Address
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                        Category
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                        Status & Date
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                        History
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                        Action
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {group.farmers.map((farmer: any, idx: number) => {
                      const isChecked = selectedFarmerIds.includes(
                        String(farmer.id)
                      );
                      return (
                        <tr key={farmer.id} className="hover:bg-slate-50">
                          <td className="border border-slate-300 px-2 py-1.5 text-center">
                            <input
                              type="checkbox"
                              checked={isChecked}
                              onChange={() =>
                                toggleFarmerCheckbox(String(farmer.id))
                              }
                              className="w-3.5 h-3.5 accent-emerald-600 cursor-pointer"
                            />
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5">
                            {idx + 1}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5 font-bold text-slate-800">
                            {farmer.farmer_name}
                            {farmer.runningCount > 1 && (
                              <span className="ml-1.5 px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 text-[10px]">
                                {farmer.runningCount} Batches
                              </span>
                            )}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5">
                            {farmer.mobile || "-"}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5">
                            {[farmer.village, farmer.upazila]
                              .filter(Boolean)
                              .join(", ")}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5 font-semibold text-slate-700">
                            {farmer.category}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5">
                            <div className="flex flex-wrap items-center gap-1.5">
                              <span
                                className={`inline-block px-2 py-0.5 rounded-full text-[11px] font-bold ${
                                  farmer.computedStatus === "Running"
                                    ? "bg-emerald-100 text-emerald-800"
                                    : farmer.computedStatus === "Active"
                                    ? "bg-blue-100 text-blue-800"
                                    : "bg-amber-100 text-amber-800"
                                }`}
                              >
                                {farmer.computedStatus}
                              </span>

                              {farmer.nourishTag && (
                                <span
                                  className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                    farmer.nourishTag === "Convert to Nourish"
                                      ? "bg-teal-100 text-teal-800 border border-teal-300"
                                      : "bg-red-100 text-red-800 border border-red-300"
                                  }`}
                                >
                                  {farmer.nourishTag}
                                  {farmer.nourishTagDate
                                    ? ` (${farmer.nourishTagDate})`
                                    : ""}
                                </span>
                              )}
                            </div>

                            {farmer.statusDate && (
                              <div className="text-[10px] text-slate-500 mt-0.5">
                                {farmer.computedStatus === "Running"
                                  ? `H.Date: ${farmer.statusDate}`
                                  : `Sale/Cull: ${farmer.statusDate}`}
                              </div>
                            )}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5 text-center">
                            <button
                              type="button"
                              onClick={() => handleOpenHistory(farmer)}
                              className="bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 px-2.5 py-1 rounded text-[11px] font-semibold"
                            >
                              History
                            </button>
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5 text-center whitespace-nowrap">
                            <button
                              type="button"
                              onClick={() => handleOpenSingleActivate(farmer)}
                              className="bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold shadow-sm"
                            >
                              {farmer.computedStatus === "Running"
                                ? "+ Add Batch"
                                : "Activate"}
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* ==================================================
          4. ACTIVATE / ADD BATCH MODAL
      ================================================== */}
      {showModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-md shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <div>
                <h3 className="text-sm sm:text-base font-bold text-slate-800">
                  {isBulkMode
                    ? `Bulk Activate Farmers (${selectedFarmerIds.length})`
                    : selectedFarmer?.runningCount > 0
                    ? `Add New Batch: ${selectedFarmer?.farmer_name} (${
                        selectedFarmer.runningCount + 1
                      })`
                    : `Activate Farmer: ${selectedFarmer?.farmer_name}`}
                </h3>
              </div>
              <button
                onClick={() => setShowModal(false)}
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
                <select
                  value={form.category}
                  onChange={(e) => {
                    const cat = e.target.value;
                    const breeds = BREEDS_BY_CATEGORY[cat] || [];
                    setForm({
                      ...form,
                      category: cat,
                      breed: breeds[0] || "",
                    });
                  }}
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white outline-none"
                >
                  {CATEGORIES.map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Hatch Date *
                </label>
                <input
                  type="date"
                  value={form.hatch_date}
                  onChange={(e) =>
                    setForm({ ...form, hatch_date: e.target.value })
                  }
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Hatchery (All Companies) *
                </label>
                <select
                  value={form.hatchery}
                  onChange={(e) =>
                    setForm({ ...form, hatchery: e.target.value })
                  }
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white outline-none"
                >
                  <option value="">Select Hatchery Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.company_name}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Breed *
                </label>
                <select
                  value={form.breed}
                  onChange={(e) => setForm({ ...form, breed: e.target.value })}
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white outline-none"
                >
                  <option value="">Select Breed</option>
                  {(BREEDS_BY_CATEGORY[form.category] || []).map((b) => (
                    <option key={b} value={b}>
                      {b}
                    </option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-2.5">
                <div>
                  <label className="block font-bold text-slate-600 mb-1">
                    Box No.
                  </label>
                  <input
                    type="text"
                    placeholder="Box No"
                    value={form.box_no}
                    onChange={(e) =>
                      setForm({ ...form, box_no: e.target.value })
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
                    value={form.t_birds}
                    onChange={(e) =>
                      setForm({ ...form, t_birds: e.target.value })
                    }
                    className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none"
                  />
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t">
                <button
                  type="button"
                  onClick={() => setShowModal(false)}
                  className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 font-semibold"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleActivateSubmit}
                  className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold shadow-sm"
                >
                  Active & Move to Running
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          5. HISTORY MODAL (WITH STANDARDS, FEED CONVERSION LOG & EXCEL EXPORT)
      ================================================== */}
      {showHistoryModal && historyFarmer && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-6xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-5">
            <div className="flex flex-wrap justify-between items-center gap-2 border-b pb-3">
              <div>
                <h3 className="text-base font-bold text-slate-800">
                  History — {historyFarmer.farmer_name} ({historyFarmer.category})
                </h3>
                <p className="text-xs text-slate-500">
                  Mobile: {historyFarmer.mobile || "-"} • Address:{" "}
                  {historyFarmer.village}, {historyFarmer.upazila}
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={exportSingleFarmerHistoryExcel}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold"
                >
                  Export History Excel
                </button>
                <button
                  onClick={() => setShowHistoryModal(false)}
                  className="px-2.5 py-1 rounded bg-slate-100 text-slate-600 font-bold"
                >
                  ✖
                </button>
              </div>
            </div>

            {/* Feed Conversion Log Section (If Farmer ever changed feed company) */}
            {(historyData.feedLogs.length > 0 || historyFarmer.nourishTag) && (
              <div className="bg-amber-50/70 border border-amber-200 rounded-xl p-3 space-y-1.5">
                <h4 className="text-xs font-bold uppercase text-amber-900">
                  Feed Conversion History
                </h4>
                {historyData.feedLogs.length > 0 ? (
                  <div className="flex flex-wrap gap-2 text-xs">
                    {historyData.feedLogs.map((log: any) => (
                      <span
                        key={log.id}
                        className="bg-white border border-amber-300 px-2.5 py-1 rounded-lg text-slate-800 font-medium"
                      >
                        <b>{log.change_date}:</b> {log.from_company || "Prev"} ➔{" "}
                        <b className="text-emerald-700">{log.to_company}</b>{" "}
                        {log.conversion_type ? `(${log.conversion_type})` : ""}
                      </span>
                    ))}
                  </div>
                ) : (
                  <div className="text-xs text-slate-700 font-semibold">
                    {historyFarmer.nourishTag} on{" "}
                    {historyFarmer.nourishTagDate || "-"}
                  </div>
                )}
              </div>
            )}

            {(historyFarmer.category || "").toLowerCase() === "layer" ? (
              /* LAYER HISTORY TABLE WITH STANDARDS */
              <div>
                <h4 className="text-xs font-bold uppercase text-indigo-800 mb-2">
                  Layer Performance History ({historyData.layer.length})
                </h4>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse border border-slate-300 text-xs">
                    <thead className="bg-indigo-50 text-indigo-950">
                      <tr>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Visit Date
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          H.Date
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          T.B
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Age
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed / Std (g/b/d)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Egg% / Std%
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Egg Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed & Type
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Remarks
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {historyData.layer.length === 0 ? (
                        <tr>
                          <td
                            colSpan={10}
                            className="border border-slate-300 py-4 text-center text-slate-400"
                          >
                            No layer performance reports found.
                          </td>
                        </tr>
                      ) : (
                        historyData.layer.map((r: any, i: number) => {
                          const v =
                            allVisitsMap[String(r.vet_visit_id)] ||
                            historyFarmer.latestVisit;
                          const hDate = v?.hatch_date || "-";
                          const tBirds = v?.t_birds || "-";
                          const ageDays = calcAgeDays(v?.hatch_date, r.visit_date);
                          const ageWks = Math.max(1, Math.floor(ageDays / 7));
                          const std = layerStds[ageWks];

                          return (
                            <tr key={i} className="hover:bg-slate-50">
                              <td className="border border-slate-300 px-2 py-1">
                                {r.visit_date || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {hDate}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                {tBirds}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                {ageWks}w
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b>{r.b_weight_g}</b> /{" "}
                                <span className="text-slate-500">
                                  {std?.weight || "-"}
                                </span>
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b>{r.feed_intake_g}</b> /{" "}
                                <span className="text-slate-500">
                                  {std?.feed_bird_day || "-"}
                                </span>
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b className="text-indigo-700">
                                  {r.egg_production}%
                                </b>{" "}
                                /{" "}
                                <span className="text-slate-500">
                                  {std?.egg_production || "-"}%
                                </span>
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b>{r.egg_weight}</b> /{" "}
                                <span className="text-slate-500">
                                  {std?.egg_weight || "-"}
                                </span>
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.feed_company || "-"}{" "}
                                {r.feed_type ? `(${r.feed_type})` : ""}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.remarks || "-"}
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
              /* BROILER / SONALI BEFORE & AFTER SALE TABLES WITH STANDARDS */
              <>
                <div>
                  <h4 className="text-xs font-bold uppercase text-emerald-800 mb-2">
                    Before Sale History ({historyData.before.length})
                  </h4>
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse border border-slate-300 text-xs">
                      <thead className="bg-slate-100">
                        <tr>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Visit Date
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            H.Date
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            T.B
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Age
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Mort.
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Wt / Std (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed (Kg)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed/B / Std (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR(w/o) / Std
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR(With M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed & Type
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Remarks
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {historyData.before.length === 0 ? (
                          <tr>
                            <td
                              colSpan={12}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No before sale reports found.
                            </td>
                          </tr>
                        ) : (
                          historyData.before.map((r: any, i: number) => {
                            const v =
                              allVisitsMap[String(r.vet_visit_id)] ||
                              historyFarmer.latestVisit;
                            const hDate = v?.hatch_date || "-";
                            const tBirds = v?.t_birds || "-";
                            const ageDays = calcAgeDays(
                              v?.hatch_date,
                              r.visit_date
                            );
                            const std = broilerStds[ageDays];

                            return (
                              <tr key={i} className="hover:bg-slate-50">
                                <td className="border border-slate-300 px-2 py-1">
                                  {r.visit_date || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {hDate}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                  {tBirds}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                  {ageDays}d
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {r.mortality_pcs}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{r.weight_gm_b}</b> /{" "}
                                  <span className="text-slate-500">
                                    {std?.weight || "-"}
                                  </span>
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {r.feed_t_kg}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{r.feed_gm_b}</b> /{" "}
                                  <span className="text-slate-500">
                                    {std?.feed_intake || "-"}
                                  </span>
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                  <b className="text-emerald-700">
                                    {r.fcr_w_out_m}
                                  </b>{" "}
                                  /{" "}
                                  <span className="text-slate-500">
                                    {std?.fcr || "-"}
                                  </span>
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                  {r.fcr_with_m}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {r.feed_company || "-"}{" "}
                                  {r.feed_type ? `(${r.feed_type})` : ""}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {r.remarks || "-"}
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div>
                  <h4 className="text-xs font-bold uppercase text-blue-800 mb-2">
                    After Sale History ({historyData.after.length})
                  </h4>
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse border border-slate-300 text-xs">
                      <thead className="bg-slate-100">
                        <tr>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Sale Date
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            H.Date
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            T.B
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Age
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Mort.
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Wt / Std (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Total Feed (Kg)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed/B / Std (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR(w/o) / Std
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR(With M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed Mill & Type
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Remarks
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {historyData.after.length === 0 ? (
                          <tr>
                            <td
                              colSpan={12}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No after sale reports found.
                            </td>
                          </tr>
                        ) : (
                          historyData.after.map((r: any, i: number) => {
                            const v =
                              allVisitsMap[String(r.vet_visit_id)] ||
                              historyFarmer.latestVisit;
                            const hDate = v?.hatch_date || "-";
                            const tBirds = Number(v?.t_birds || 0);
                            const ageDays = calcAgeDays(
                              v?.hatch_date,
                              r.visit_date
                            );
                            const std = broilerStds[ageDays];
                            const feedGmB =
                              r.feed_gm_b ||
                              (tBirds > 0
                                ? (
                                    (Number(r.feed_consumption_kg || 0) *
                                      1000) /
                                    tBirds
                                  ).toFixed(1)
                                : "-");

                            return (
                              <tr key={i} className="hover:bg-slate-50">
                                <td className="border border-slate-300 px-2 py-1">
                                  {r.visit_date || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {hDate}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                  {tBirds || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                  {ageDays}d
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {r.mortality}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{r.weight}</b> /{" "}
                                  <span className="text-slate-500">
                                    {std?.weight || "-"}
                                  </span>
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {r.feed_consumption_kg}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{feedGmB}</b> /{" "}
                                  <span className="text-slate-500">
                                    {std?.feed_intake || "-"}
                                  </span>
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                  <b className="text-blue-700">
                                    {r.fcr_w_out_m}
                                  </b>{" "}
                                  /{" "}
                                  <span className="text-slate-500">
                                    {std?.fcr || "-"}
                                  </span>
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                  {r.fcr_with_m}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {r.feed_mill || "-"}{" "}
                                  {r.feed ? `(${r.feed})` : ""}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {r.remarks || "-"}
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
    </div>
  );
}