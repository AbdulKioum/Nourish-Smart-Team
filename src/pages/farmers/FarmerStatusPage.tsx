import { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import {
  ChevronDown,
  ChevronUp,
  History,
  Plus,
  RotateCcw,
  Search,
  X,
  CheckSquare,
} from "lucide-react";

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
  const [loading, setLoading] = useState(true);

  // Multiple Selection State
  const [selectedFarmerIds, setSelectedFarmerIds] = useState<string[]>([]);
  const [isBulkMode, setIsBulkMode] = useState(false);

  // Activate / Add Batch Modal State
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

  // History Modal State
  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [historyFarmer, setHistoryFarmer] = useState<any>(null);
  const [historyData, setHistoryData] = useState<{
    before: any[];
    after: any[];
    layer: any[];
  }>({ before: [], after: [], layer: [] });

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
      ] = await Promise.all([
        supabase.from("companies").select("*").order("company_name"),
        supabase.from("agents").select("*").order("agent_name"),
        supabase.from("farmers").select("*").order("farmer_name"),
        supabase
          .from("vet_visits")
          .select("*")
          .order("created_at", { ascending: false }),
      ]);

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
    if (selectedFarmerIds.length === 0) {
      alert("Please select at least one farmer using the checkboxes.");
      return;
    }
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
        // If farmer already has a running batch, label with 1, 2, etc.
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

        // If a 2nd batch is added while 1st batch is still unlabeled, label 1st batch as (1)
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

      alert(
        isBulkMode
          ? `${targets.length} Farmers activated & moved to Running Vet Visits!`
          : `Farmer activated & moved to Running Vet Visits!`
      );
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

    const [{ data: bData }, { data: aData }, { data: lData }] =
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
      ]);

    setHistoryData({
      before: bData || [],
      after: aData || [],
      layer: lData || [],
    });
  }

  const availableAgentsForDropdown =
    selectedCompany === "all"
      ? agents
      : agents.filter((a) => String(a.company_id) === String(selectedCompany));

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-4">
      {/* Header + Bulk Activate Button */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-4 py-2.5 shadow-sm">
        <div className="flex items-center gap-3">
          <h1 className="text-base sm:text-lg font-bold text-slate-800">
            Farmer Status Management
          </h1>
          <span className="text-xs font-semibold bg-slate-100 border border-slate-200 px-2.5 py-0.5 rounded-lg text-slate-700">
            Showing: <strong>{tableFilteredFarmers.length}</strong>
          </span>
        </div>

        {selectedFarmerIds.length > 0 && (
          <button
            type="button"
            onClick={handleOpenBulkActivate}
            className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold shadow-sm transition"
          >
            <CheckSquare size={14} />
            Activate Selected ({selectedFarmerIds.length})
          </button>
        )}
      </div>

      {/* ==================================================
          1. COMPACT SUMMARY CARDS WITH "DETAILS" TOGGLE
      ================================================== */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2.5">
        {summaryCards.map((card) => {
          const isSelected = selectedStatusFilter === card.key;
          const isExpanded = expandedCard === card.key;

          return (
            <div
              key={card.key}
              className={`rounded-xl border p-2.5 shadow-sm transition flex flex-col justify-between ${
                card.color
              } ${isSelected ? "ring-2 ring-emerald-600" : ""}`}
            >
              <div
                onClick={() => setSelectedStatusFilter(card.key)}
                className="cursor-pointer flex items-center justify-between gap-1"
              >
                <span className="text-xs font-bold leading-tight">
                  {card.title}
                </span>
                <span
                  className={`px-2 py-0.5 rounded-full text-xs font-extrabold ${card.badgeBg}`}
                >
                  {card.data.total}
                </span>
              </div>

              <div className="mt-2 pt-1.5 border-t border-black/10 flex items-center justify-between text-[11px]">
                <button
                  type="button"
                  onClick={() => setSelectedStatusFilter(card.key)}
                  className="hover:underline font-medium opacity-80"
                >
                  {isSelected ? "✓ Filtered" : "Filter"}
                </button>
                <button
                  type="button"
                  onClick={() =>
                    setExpandedCard(isExpanded ? null : card.key)
                  }
                  className="inline-flex items-center gap-0.5 font-bold hover:underline"
                >
                  Details{" "}
                  {isExpanded ? (
                    <ChevronUp size={12} />
                  ) : (
                    <ChevronDown size={12} />
                  )}
                </button>
              </div>

              {isExpanded && (
                <div className="mt-2 pt-2 border-t border-black/10 space-y-1 text-[11px] bg-white/90 rounded-lg p-2 text-slate-800">
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
          <div className="col-span-2 sm:col-span-1 relative">
            <Search
              size={13}
              className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Search farmer, mobile..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full rounded-lg border border-slate-300 pl-8 pr-2.5 py-1.5 text-xs outline-none focus:border-emerald-500"
            />
          </div>

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
            className="inline-flex items-center justify-center gap-1 rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 transition"
          >
            <RotateCcw size={13} /> Reset
          </button>
        </div>
      </div>

      {/* ==================================================
          3. AGENT-WISE MINIMAL BORDERED FARMER TABLE
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
                      <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                        Nourish Conversion
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
                                {farmer.runningCount} Batches Running
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
                            <div className="flex flex-col gap-0.5">
                              <span
                                className={`inline-block w-fit px-2 py-0.5 rounded-full text-[11px] font-bold ${
                                  farmer.computedStatus === "Running"
                                    ? "bg-emerald-100 text-emerald-800"
                                    : farmer.computedStatus === "Active"
                                    ? "bg-blue-100 text-blue-800"
                                    : "bg-amber-100 text-amber-800"
                                }`}
                              >
                                {farmer.computedStatus}
                              </span>
                              {farmer.statusDate && (
                                <span className="text-[10px] text-slate-500">
                                  {farmer.computedStatus === "Running"
                                    ? `H.Date: ${farmer.statusDate}`
                                    : `Sale/Cull: ${farmer.statusDate}`}
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5">
                            {farmer.nourishTag ? (
                              <div>
                                <span
                                  className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold ${
                                    farmer.nourishTag === "Convert to Nourish"
                                      ? "bg-teal-100 text-teal-800"
                                      : "bg-red-100 text-red-800"
                                  }`}
                                >
                                  {farmer.nourishTag}
                                </span>
                                {farmer.nourishTagDate && (
                                  <div className="text-[10px] text-slate-500">
                                    Date: {farmer.nourishTagDate}
                                  </div>
                                )}
                              </div>
                            ) : (
                              <span className="text-slate-400">-</span>
                            )}
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5 text-center">
                            <button
                              type="button"
                              onClick={() => handleOpenHistory(farmer)}
                              className="inline-flex items-center gap-1 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 px-2 py-1 rounded text-[11px] font-semibold"
                            >
                              <History size={11} /> History
                            </button>
                          </td>
                          <td className="border border-slate-300 px-2.5 py-1.5 text-center whitespace-nowrap">
                            <button
                              type="button"
                              onClick={() => handleOpenSingleActivate(farmer)}
                              className="inline-flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold shadow-sm"
                            >
                              <Plus size={11} />
                              {farmer.computedStatus === "Running"
                                ? "Add Batch"
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
          4. ACTIVATE / ADD BATCH MODAL (SINGLE OR BULK)
      ================================================== */}
      {showModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-md shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <div>
                <h3 className="text-sm sm:text-base font-bold text-slate-800">
                  {isBulkMode
                    ? `Bulk Activate Farmers (${selectedFarmerIds.length} Selected)`
                    : selectedFarmer?.runningCount > 0
                    ? `Add New Batch: ${selectedFarmer?.farmer_name} (${
                        selectedFarmer.runningCount + 1
                      })`
                    : `Activate Farmer: ${selectedFarmer?.farmer_name}`}
                </h3>
                {!isBulkMode && selectedFarmer && (
                  <p className="text-[11px] text-slate-500">
                    Agent: {selectedFarmer.agentName} (
                    {selectedFarmer.companyName})
                  </p>
                )}
              </div>
              <button onClick={() => setShowModal(false)}>
                <X size={18} />
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
          5. HISTORY MODAL (CATEGORY-SPECIFIC CLEAN BORDERED TABLES)
      ================================================== */}
      {showHistoryModal && historyFarmer && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-5xl max-h-[90vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-5">
            <div className="flex justify-between items-center border-b pb-3">
              <div>
                <h3 className="text-base font-bold text-slate-800">
                  Visit & Report History — {historyFarmer.farmer_name} (
                  {historyFarmer.category})
                </h3>
                <p className="text-xs text-slate-500">
                  Mobile: {historyFarmer.mobile || "-"} • Address:{" "}
                  {historyFarmer.village}, {historyFarmer.upazila}
                </p>
              </div>
              <button onClick={() => setShowHistoryModal(false)}>
                <X size={18} />
              </button>
            </div>

            {(historyFarmer.category || "").toLowerCase() === "layer" ? (
              /* ONLY LAYER PERFORMANCE HISTORY FOR LAYER FARMERS */
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
                          B. Weight (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed Intake (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Egg Prod (%)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Egg Wt (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed Company
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Disease / Vaccine
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
                            colSpan={8}
                            className="border border-slate-300 py-4 text-center text-slate-400"
                          >
                            No layer performance reports found.
                          </td>
                        </tr>
                      ) : (
                        historyData.layer.map((r: any, i: number) => (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="border border-slate-300 px-2 py-1">
                              {r.visit_date || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.b_weight_g}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.feed_intake_g}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold text-indigo-700">
                              {r.egg_production}%
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.egg_weight}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.feed_company || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {[r.disease, r.vaccine]
                                .filter(Boolean)
                                .join(" / ") || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.remarks || "-"}
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              /* ONLY BEFORE & AFTER SALE TABLES FOR BROILER / SONALI FARMERS */
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
                            Mortality
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Weight (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed (Kg)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed (g/b)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR (w/o M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR (With M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed Company
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed Type
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
                              colSpan={10}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No before sale reports found.
                            </td>
                          </tr>
                        ) : (
                          historyData.before.map((r: any, i: number) => (
                            <tr key={i} className="hover:bg-slate-50">
                              <td className="border border-slate-300 px-2 py-1">
                                {r.visit_date || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {r.mortality_pcs}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {r.weight_gm_b}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {r.feed_t_kg}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {r.feed_gm_b}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                {r.fcr_w_out_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                {r.fcr_with_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.feed_company || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.feed_type || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.remarks || "-"}
                              </td>
                            </tr>
                          ))
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
                            Sale / Visit Date
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Mortality
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Avg Weight (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Total Feed (Kg)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR (w/o M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR (With M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed Mill
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed
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
                              colSpan={9}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No after sale reports found.
                            </td>
                          </tr>
                        ) : (
                          historyData.after.map((r: any, i: number) => (
                            <tr key={i} className="hover:bg-slate-50">
                              <td className="border border-slate-300 px-2 py-1">
                                {r.visit_date || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {r.mortality}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {r.weight}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {r.feed_consumption_kg}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                {r.fcr_w_out_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-semibold">
                                {r.fcr_with_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.feed_mill || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.feed || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {r.remarks || "-"}
                              </td>
                            </tr>
                          ))
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