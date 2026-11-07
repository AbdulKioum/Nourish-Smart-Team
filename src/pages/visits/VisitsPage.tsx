import React, { useEffect, useRef, useState } from "react";
import { useLocation } from "react-router-dom";
import { supabase } from "../../lib/supabase";
import { useAuth } from "../../context/AuthContext";
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

const STOCK_KEY_BY_CATEGORY: Record<string, string> = {
  Broiler: "broiler",
  Layer: "layer",
  Sonali: "sonali",
  Duck: "duck",
  Quail: "quail",
};
const POULTRY_STOCK_KEYS = ["broiler", "layer", "sonali", "duck", "quail"];
const STATUS_RANK: Record<string, number> = { Running: 0, Active: 1, Inactive: 2 };

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

// Batch tag shown under the farmer name (e.g. "Sonali Batch (1)")
function batchTagOf(label?: string, farmerName?: string): string {
  if (!label) return "";
  if (label === farmerName) return "";
  if (farmerName && label.startsWith(farmerName)) {
    const rest = label.slice(farmerName.length).trim();
    return rest;
  }
  return label;
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
  // Broiler Exact Schedule
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
  // Sonali Exact Schedule
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
  // Layer Schedule
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

export default function VetVisitsPage() {
  const { role } = useAuth();
  const isSuperAdmin = role === "super_admin";

  const [currentUserProfile, setCurrentUserProfile] = useState({
    fullName: "Officer",
    designation: "Customer Service Officer",
    region: "Mymensingh Zone",
  });

  const [companies, setCompanies] = useState<any[]>([]);
  const [allAgents, setAllAgents] = useState<any[]>([]);
  const [selectedCompany, setSelectedCompany] = useState<string>("all");
  const [selectedCategoryFilter, setSelectedCategoryFilter] =
    useState<string>("all");

  const [_visitsData, setVisitsData] = useState<any[]>([]);
  const location = useLocation();
  const companyInitRef = useRef(false);
  const [statusFarmers, setStatusFarmers] = useState<any[]>([]);
  const [selectedAgent, setSelectedAgent] = useState<string>("all");
  const [selectedStatusFilter, setSelectedStatusFilter] =
    useState<StatusTab>("all");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [expandedCard, setExpandedCard] = useState<string | null>(null);
  const [focusVisitId, setFocusVisitId] = useState<string | null>(null);

  // Activate / Run Farm / Add Batch
  const [showActivateModal, setShowActivateModal] = useState(false);
  const [activateFarmer, setActivateFarmer] = useState<any>(null);
  const [activateBulk, setActivateBulk] = useState(false);
  const [selectedFarmerIds, setSelectedFarmerIds] = useState<string[]>([]);
  const [activateMode, setActivateMode] = useState<"add" | "replace">("add");
  const [activateForm, setActivateForm] = useState({
    category: "Broiler",
    hatch_date: new Date().toISOString().split("T")[0],
    hatchery: "",
    breed: "",
    box_no: "",
    t_birds: "",
    batch_mark: "",
  });

  // Edit a saved report (from History)
  const [editReport, setEditReport] = useState<{
    table: string;
    id: any;
    form: Record<string, any>;
  } | null>(null);
  const [vaccines, setVaccines] = useState<any[]>([]);
  const [broilerStds, setBroilerStds] = useState<Record<number, any>>({});
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

  // Quick "+" Log Given Vaccine Modal for a Specific Running Batch
  const [showLogVaccineModal, setShowLogVaccineModal] = useState(false);
  const [activeVaccineVisit, setActiveVaccineVisit] = useState<any>(null);
  const [logVaccineForm, setLogVaccineForm] = useState({
    vaccine_name: "",
    vaccine_date: new Date().toISOString().split("T")[0],
  });

  // Edit Batch Modal
  const [showEditModal, setShowEditModal] = useState(false);
  const [editingVisit, setEditingVisit] = useState<any>(null);
  const [editForm, setEditForm] = useState({
    hatch_date: "",
    hatchery: "",
    breed: "",
    box_no: "",
    t_birds: "",
    category: "Broiler",
  });

  // History Modal
  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [selectedHistoryVisit, setSelectedHistoryVisit] = useState<any>(null);
  const [farmerHistory, setFarmerHistory] = useState<{
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

  const [transferAgentId, setTransferAgentId] = useState<string>("");

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

  // Layer Form (Vaccine removed from here as requested; tracked via "+" button in table)
  const [layerForm, setLayerForm] = useState({
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
    loadVetVisits();
  }, []);

  // Default company = Nourish (unless we arrived from a notification)
  useEffect(() => {
    if (companyInitRef.current || companies.length === 0) return;
    companyInitRef.current = true;
    const hasFocus = new URLSearchParams(location.search).get("focus");
    if (hasFocus) return;
    const nourish = companies.find((c) =>
      c.company_name?.toLowerCase().includes("nourish")
    );
    if (nourish) setSelectedCompany(String(nourish.id));
  }, [companies]);

  // Coming from the notification "Sold" button: open that farmer's After Sale option
  useEffect(() => {
    const fid = new URLSearchParams(location.search).get("focus");
    if (!fid) return;
    setFocusVisitId(fid);
    setSelectedCompany("all");
    setSelectedAgent("all");
    setSelectedStatusFilter("all");
    setSelectedCategoryFilter("all");
    setSearchQuery("");
  }, [location.search]);

  useEffect(() => {
    if (!focusVisitId || _loading) return;
    const t = setTimeout(() => {
      const el =
        document.getElementById("after-sale-btn-" + focusVisitId) ||
        document.getElementById("visit-row-" + focusVisitId);
      el?.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 400);
    return () => clearTimeout(t);
  }, [focusVisitId, statusFarmers, _loading]);

  async function loadInitialData() {
    setLoading(true);
    const {
      data: { user },
    } = await supabase.auth.getUser();

    if (user) {
      const { data: uProf } = await supabase
        .from("users")
        .select("full_name, email, designation, region")
        .eq("id", user.id)
        .single();

      if (uProf) {
        setCurrentUserProfile({
          fullName: uProf.full_name || uProf.email?.split("@")[0] || "Officer",
          designation:
            uProf.designation || "Customer Service Officer (Poultry)",
          region: uProf.region || "Mymensingh Zone",
        });
      }
    }

    const [
      { data: compData },
      { data: agData },
      { data: vacData },
      { data: bStdData },
      { data: lStdData },
    ] = await Promise.all([
      supabase.from("companies").select("*").order("company_name"),
      supabase.from("agents").select("*").order("agent_name"),
      supabase
        .from("vaccines")
        .select("*")
        .order("age_days", { ascending: true }),
      supabase.from("broiler_standards").select("*"),
      supabase.from("layer_standards").select("*"),
    ]);

    const bMap: Record<number, any> = {};
    (bStdData || []).forEach((s: any) => (bMap[Number(s.age_days)] = s));
    setBroilerStds(bMap);

    const lMap: Record<string, any> = {};
    (lStdData || []).forEach((s: any) => {
      const breedKey = String(s.breed || "shaver star cross 579 brown")
        .trim()
        .toLowerCase();
      lMap[`${breedKey}_${Number(s.age_weeks)}`] = s;
      // Fallback by week number if old schema was used
      if (!lMap[`any_${Number(s.age_weeks)}`]) {
        lMap[`any_${Number(s.age_weeks)}`] = s;
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
    setAllAgents(agData || []);

    const nourish = comps.find((c) =>
      c.company_name?.toLowerCase().includes("nourish")
    );
    if (nourish) {
      setSelectedCompany(String(nourish.id));
    } else if (comps.length > 0) {
      setSelectedCompany(String(comps[0].id));
    }
    setLoading(false);
  }

  async function loadVaccinesList() {
    const { data } = await supabase
      .from("vaccines")
      .select("*")
      .order("age_days", { ascending: true });
    setVaccines(data || []);
  }

  async function loadVetVisits() {
    setLoading(true);
    try {
      const [
        { data: visits, error: visitErr },
        { data: agentsData },
        { data: farmersData },
      ] = await Promise.all([
        supabase
          .from("vet_visits")
          .select("*")
          .order("created_at", { ascending: false }),
        supabase.from("agents").select("*").order("agent_name"),
        supabase.from("farmers").select("*").order("farmer_name"),
      ]);
      if (visitErr) throw visitErr;

      const allVisits = visits || [];
      const agentsList = agentsData || [];
      const farmersList = farmersData || [];

      const enrich = (v: any) => ({
        ...v,
        agents:
          agentsList.find((a: any) => String(a.id) === String(v.agent_id)) || {
            agent_name: "Unknown Agent",
          },
        farmers:
          farmersList.find((f: any) => String(f.id) === String(v.farmer_id)) || {
            farmer_name: "Unknown Farmer",
            mobile: "-",
            village: "-",
            upazila: "-",
          },
      });

      const processed = farmersList.map((f: any) => {
        const fv = allVisits.filter(
          (v: any) => String(v.farmer_id) === String(f.id)
        );
        const runningVisits = fv
          .filter((v: any) => v.status === "Activated" || v.status === "Running")
          .sort((a: any, b: any) =>
            String(a.created_at || "").localeCompare(String(b.created_at || ""))
          )
          .map(enrich);
        const closedVisits = fv.filter(
          (v: any) => v.status === "Sold" || v.status === "Culled"
        );
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
        return {
          ...f,
          computedStatus,
          statusDate,
          daysSinceSale,
          closedType: latestClosed?.status || "",
          closedCategory: latestClosed?.category || "",
          category: runningVisits[0]?.category || normalizeCategory(f),
          runningVisits,
          runningCount: runningVisits.length,
          latestClosed,
          agentName: agentObj?.agent_name || "Unassigned Agent",
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
      return `${totalWeeks} Wks (${diffDays}d)`;
    }
    return `${diffDays} Days`;
  }

  function getLayerStdForVisit(breedName?: string, ageWks?: number) {
    const w = Math.min(100, Math.max(1, ageWks || 1));
    const isBovans = String(breedName || "").toLowerCase().includes("bovans");
    const targetBreedKey = isBovans
      ? "bovans white"
      : "shaver star cross 579 brown";
    return (
      layerStds[`${targetBreedKey}_${w}`] || layerStds[`any_${w}`] || null
    );
  }

  // 1. Updated getVaccineInfoForVisit (Without Due Today Alert)
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

  // Open Quick "+" Vaccine Modal for a Batch
  function handleOpenLogVaccineModal(visit: any) {
    const info = getVaccineInfoForVisit(visit);
    const defaultVacName =
      info.upcoming?.vaccine_name ||
      visit.last_vaccine_name ||
      "";

    setActiveVaccineVisit(visit);
    setLogVaccineForm({
      vaccine_name: defaultVacName,
      vaccine_date: new Date().toISOString().split("T")[0],
    });
    setShowLogVaccineModal(true);
  }

  async function handleSaveLoggedVaccine(e: React.FormEvent) {
    e.preventDefault();
    if (!activeVaccineVisit || !logVaccineForm.vaccine_name.trim()) return;

    try {
      const { error } = await supabase
        .from("vet_visits")
        .update({
          last_vaccine_name: logVaccineForm.vaccine_name.trim(),
          last_vaccine_date: logVaccineForm.vaccine_date,
        })
        .eq("id", activeVaccineVisit.id);

      if (error) throw error;

      setShowLogVaccineModal(false);
      setActiveVaccineVisit(null);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save vaccine log.");
    }
  }

  // Corrected Before Sale Auto-Calculations:
  // Without Mortality = Lower FCR (Total Birds weight)
  // With Mortality = Higher FCR (Live Birds weight after subtracting mortality)
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

  // Corrected After Sale Auto-Calculations:
  // Without Mortality = Lower FCR (Total Birds weight)
  // With Mortality = Higher FCR (Live Birds weight after subtracting mortality)
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

  // Layer Dual Feed Intake Calculation (Total Kg <-> gm/bird/day)
  function handleLayerFeedChange(
    mode: "total_kg" | "gm_bird",
    rawValue: string
  ) {
    const tBirds = Number(activeReportVisit?.t_birds) || 0;
    const num = Number(rawValue);

    if (mode === "total_kg") {
      const gmPerBird =
        tBirds > 0 && num > 0 ? ((num * 1000) / tBirds).toFixed(1) : "";
      setLayerForm((prev) => ({
        ...prev,
        feed_total_kg: rawValue,
        feed_intake_g: gmPerBird,
      }));
    } else {
      const totalKg =
        tBirds > 0 && num > 0 ? ((num * tBirds) / 1000).toFixed(2) : "";
      setLayerForm((prev) => ({
        ...prev,
        feed_intake_g: rawValue,
        feed_total_kg: totalKg,
      }));
    }
  }

  // Generate Broiler / Sonali WhatsApp Observation Message
  function generateBroilerWhatsAppText(visit: any, rep: any, vDate: string) {
    const farmer = visit.farmers || {};
    const agent = visit.agents || {};
    const ageDays = getAgeDays(visit.hatch_date, vDate);
    const std = broilerStds[ageDays];

    const actualWt = Number(rep.weight_gm_b || 0);
    const stdWt = Number(std?.weight || 0);
    const wtDiff = actualWt - stdWt;
    const wtDiffText =
      stdWt > 0
        ? `${Math.abs(wtDiff).toFixed(0)}gm ${wtDiff >= 0 ? "more" : "less"}`
        : "N/A";

    const actualFeed = Number(rep.feed_gm_b || 0);
    const stdFeed = Number(std?.feed_intake || 0);
    const feedDiff = actualFeed - stdFeed;
    const feedDiffText =
      stdFeed > 0
        ? `${Math.abs(feedDiff).toFixed(0)}gm ${
            feedDiff >= 0 ? "more" : "less"
          }`
        : "N/A";

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

    const msg = `Assalamu Alaikum sir.
*Broiler Farm visit Observation Report*
Visit Date: ${formatDmy(vDate)}

Agent: ${agent.agent_name || "-"}
Farmer Name - ${visit.batch_label || farmer.farmer_name || "-"}
Farm Address - ${[farmer.village, farmer.upazila].filter(Boolean).join(", ")}
--------------------------------------------
Hatchery - ${visit.hatchery || "-"}
Breed - ${visit.breed || "-"}
H.Date: ${formatDmy(visit.hatch_date)}
--------------------------------------------
Total Bird (Pcs) - ${visit.t_birds || 0}
Age (Day) - ${ageDays} days
Mortality (Pcs) - ${rep.mortality_pcs || 0} pcs

Body Weight - ${actualWt} gm/bird.
Std - ${stdWt || "-"} gm/bird.

Feed Intake - ${actualFeed} gm/bird.
Std - ${stdFeed || "-"} gm/bird.

FCR (without mortality): ${rep.fcr_w_out_m || "-"}
Std - ${std?.fcr || "-"}
FCR (with mortality): ${rep.fcr_with_m || "-"}
--------------------------------------------
Feed - ${rep.feed_company || "-"}
Feed Type - ${rep.feed_type || "-"}
MFG: ${rep.pro_date ? formatDmy(rep.pro_date) : "-"}
Batch: ${rep.batch_no || "-"}
${findingsBlock}
*Body weight (${wtDiffText}) than standard and Feed intake (${feedDiffText}) than standard*

Thanks.
${currentUserProfile.fullName}
${currentUserProfile.designation}
${currentUserProfile.region}`;

    setCopiedWa(false);
    setWhatsappPopupText(msg);
  }

  // Generate Layer WhatsApp Observation Message
  function generateLayerWhatsAppText(visit: any, rep: any) {
    const farmer = visit.farmers || {};
    const agent = visit.agents || {};
    const ageWks = getAgeWeeks(
      visit.hatch_date,
      rep.visit_date || reportVisitDate
    );

    const remarksLine =
      rep.remarks && String(rep.remarks).trim() !== ""
        ? `\nRemarks: ${rep.remarks.trim()}\n`
        : "";

    const msg = `Assalamualaikum Sir,
*Layer farm visit report*

Agent Name: ${agent.agent_name || "-"}
Farmer: ${visit.batch_label || farmer.farmer_name || "-"}
Address: ${[farmer.village, farmer.upazila].filter(Boolean).join(", ")}

Hatchery: ${rep.hatchery || visit.hatchery || "-"}
Breed: ${rep.breed || visit.breed || "-"}

T.B.: ${visit.t_birds || 0}
Age: ${ageWks} weeks
Weight: ${rep.b_weight_g || 0} gm
Feed Intake(gm/birds): ${rep.feed_intake_g || 0} gm
Eggs Production: ${rep.egg_production || 0}%
Egg weight: ${rep.egg_weight || 0} gm

Feed: ${rep.feed_company || "-"}
Type: ${rep.feed_type || "-"}
MFG: ${rep.pro_date ? formatDmy(rep.pro_date) : "-"}
Batch No.: ${rep.batch_no || "-"}
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
      await supabase.from("farmer_feed_logs").insert([
        {
          farmer_id: String(visit.farmer_id),
          from_company: prevFeedName,
          to_company: chosenFeedCompanyName,
          conversion_type: convType || "Feed Switch",
          change_date: visitDateStr,
        },
      ]);
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
    setActiveReportVisit(visit);
    setReportVisitDate(new Date().toISOString().split("T")[0]);
    setTransferAgentId("");

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

  async function handleBeforeSaleSubmit() {
    if (!activeReportVisit) return;
    try {
      await processFeedCompanyChangeIfNeeded(
        activeReportVisit,
        beforeForm.feed_company,
        transferAgentId,
        reportVisitDate
      );

      const { error } = await supabase.from("before_sale_reports").insert([
        {
          vet_visit_id: activeReportVisit.id,
          farmer_id: String(activeReportVisit.farmer_id),
          agent_id: String(transferAgentId || activeReportVisit.agent_id),
          company_id: String(activeReportVisit.company_id),
          visit_date: reportVisitDate,
          mortality_pcs: Number(beforeForm.mortality_pcs) || 0,
          weight_gm_b: Number(beforeForm.weight_gm_b) || 0,
          feed_t_kg: Number(beforeForm.feed_t_kg) || 0,
          feed_gm_b: Number(beforeForm.feed_gm_b) || 0,
          fcr_w_out_m: Number(beforeForm.fcr_w_out_m) || 0,
          fcr_with_m: Number(beforeForm.fcr_with_m) || 0,
          uneven: beforeForm.uneven,
          indigestion: beforeForm.indigestion,
          orange_feces: beforeForm.orange_feces,
          watery_feces: beforeForm.watery_feces,
          feeder: beforeForm.feeder,
          drinker: beforeForm.drinker,
          floor_space: beforeForm.floor_space,
          litter_condition: beforeForm.litter_condition,
          ventilation: beforeForm.ventilation,
          feed_scattering: beforeForm.feed_scattering,
          feed_hardness: beforeForm.feed_hardness,
          feed_company: beforeForm.feed_company,
          feed_type: beforeForm.feed_type,
          pro_date: beforeForm.pro_date || null,
          batch_no: beforeForm.batch_no,
          remarks: beforeForm.remarks,
        },
      ]);

      if (error) throw error;
      alert("Before Sale Report saved!");
      setShowBeforeSaleModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save report");
    }
  }

  async function handleAfterSaleSoldSubmit() {
    if (!activeReportVisit) return;
    try {
      await processFeedCompanyChangeIfNeeded(
        activeReportVisit,
        afterForm.feed_mill,
        transferAgentId,
        reportVisitDate
      );

      const payload: any = {
        vet_visit_id: activeReportVisit.id,
        farmer_id: String(activeReportVisit.farmer_id),
        agent_id: String(transferAgentId || activeReportVisit.agent_id),
        company_id: String(activeReportVisit.company_id),
        visit_date: reportVisitDate,
        mortality: Number(afterForm.mortality) || 0,
        weight: Number(afterForm.weight) || 0,
        feed_consumption_kg: Number(afterForm.feed_consumption_kg) || 0,
        feed_gm_b: Number(afterForm.feed_gm_b) || 0,
        feed: afterForm.feed,
        feed_mill: afterForm.feed_mill,
        fcr_w_out_m: Number(afterForm.fcr_w_out_m) || 0,
        fcr_with_m: Number(afterForm.fcr_with_m) || 0,
        remarks: afterForm.remarks,
      };

      const { error } = await supabase
        .from("after_sale_reports")
        .insert([payload]);

      if (error) {
        delete payload.feed_gm_b;
        const { error: retryErr } = await supabase
          .from("after_sale_reports")
          .insert([payload]);
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
      setShowAfterSaleModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save after sale report");
    }
  }

  async function handleLayerReportSubmit(markAsCulled = false) {
    if (!activeReportVisit) return;
    try {
      await processFeedCompanyChangeIfNeeded(
        activeReportVisit,
        layerForm.feed_company,
        transferAgentId,
        reportVisitDate
      );

      const payload: any = {
        vet_visit_id: activeReportVisit.id,
        farmer_id: String(activeReportVisit.farmer_id),
        agent_id: String(transferAgentId || activeReportVisit.agent_id),
        company_id: String(activeReportVisit.company_id),
        visit_date: reportVisitDate,
        b_weight_g: Number(layerForm.b_weight_g) || 0,
        feed_total_kg: Number(layerForm.feed_total_kg) || 0,
        feed_intake_g: Number(layerForm.feed_intake_g) || 0,
        egg_production: Number(layerForm.egg_production) || 0,
        egg_weight: Number(layerForm.egg_weight) || 0,
        feed_company: layerForm.feed_company,
        feed_type: layerForm.feed_type,
        pro_date: layerForm.pro_date || null,
        batch_no: layerForm.batch_no,
        hatchery: layerForm.hatchery,
        breed: layerForm.breed,
        disease: layerForm.disease,
        vaccine: activeReportVisit.last_vaccine_name || "",
        remarks: layerForm.remarks,
      };

      const { error } = await supabase.from("layer_reports").insert([payload]);

      if (error) {
        delete payload.feed_total_kg;
        const { error: retryErr } = await supabase
          .from("layer_reports")
          .insert([payload]);
        if (retryErr) throw retryErr;
      }

      if (markAsCulled) {
        await supabase
          .from("vet_visits")
          .update({ status: "Culled", closed_date: reportVisitDate })
          .eq("id", activeReportVisit.id);

        await supabase
          .from("farmers")
          .update({
            status: "Active",
            last_sale_date: reportVisitDate,
          })
          .eq("id", activeReportVisit.farmer_id);

        alert("Layer Batch marked as CULLED & moved to Active status!");
      } else {
        alert("Layer Report saved!");
      }

      setShowLayerModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save layer report");
    }
  }

  async function handleAddVaccine(e: React.FormEvent) {
    e.preventDefault();
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
  }

  async function handleDeleteVaccine(id: number) {
    if (!isSuperAdmin) return;
    await supabase.from("vaccines").delete().eq("id", id);
    await loadVaccinesList();
  }

  function handleEditClick(visit: any) {
    setEditingVisit(visit);
    setEditForm({
      hatch_date: visit.hatch_date || "",
      hatchery: visit.hatchery || "",
      breed: visit.breed || "",
      box_no: visit.box_no || "",
      t_birds: visit.t_birds || "",
      category: visit.category || "Broiler",
    });
    setShowEditModal(true);
  }

  async function handleUpdateSubmit() {
    if (!editingVisit) return;
    try {
      const { error } = await supabase
        .from("vet_visits")
        .update({
          hatch_date: editForm.hatch_date,
          hatchery: editForm.hatchery,
          breed: editForm.breed,
          box_no: editForm.box_no,
          t_birds: Number(editForm.t_birds),
          category: editForm.category,
        })
        .eq("id", editingVisit.id);

      if (error) throw error;
      alert("Vet Visit details updated!");
      setShowEditModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to update");
    }
  }

  async function handleHistoryClick(visit: any) {
    setSelectedHistoryVisit(visit);
    setShowHistoryModal(true);

    const isLayer = (visit.category || "").toLowerCase() === "layer";
    const { data: fLogs } = await supabase
      .from("farmer_feed_logs")
      .select("*")
      .eq("farmer_id", String(visit.farmer_id))
      .order("change_date", { ascending: false });

    if (isLayer) {
      const { data: lData } = await supabase
        .from("layer_reports")
        .select("*")
        .eq("vet_visit_id", String(visit.id))
        .order("visit_date", { ascending: false });

      setFarmerHistory({
        before: [],
        after: [],
        layer: lData || [],
        feedLogs: fLogs || [],
      });
    } else {
      const [{ data: bData }, { data: aData }] = await Promise.all([
        supabase
          .from("before_sale_reports")
          .select("*")
          .eq("vet_visit_id", String(visit.id))
          .order("visit_date", { ascending: false }),
        supabase
          .from("after_sale_reports")
          .select("*")
          .eq("vet_visit_id", String(visit.id))
          .order("visit_date", { ascending: false }),
      ]);

      setFarmerHistory({
        before: bData || [],
        after: aData || [],
        layer: [],
        feedLogs: fLogs || [],
      });
    }
  }

  async function exportVisitHistoryExcel() {
    if (!selectedHistoryVisit) return;
    const workbook = new ExcelJS.Workbook();
    const isLayer =
      (selectedHistoryVisit.category || "").toLowerCase() === "layer";

    if (isLayer) {
      const s = workbook.addWorksheet("Layer History");
      s.addRow([
        "Visit Date",
        "H.Date",
        "T.Birds",
        "Age (Wks)",
        "Wt (g)",
        "Std Wt",
        "Feed (g/b)",
        "Std Feed",
        "Egg%",
        "Std Egg%",
        "Egg Wt",
        "Std Egg Wt",
        "Feed",
        "Type",
        "Remarks",
      ]);
      farmerHistory.layer.forEach((r: any) => {
        const wks = getAgeWeeks(selectedHistoryVisit.hatch_date, r.visit_date);
        const std = getLayerStdForVisit(selectedHistoryVisit.breed, wks);
        s.addRow([
          r.visit_date,
          selectedHistoryVisit.hatch_date,
          selectedHistoryVisit.t_birds,
          wks,
          r.b_weight_g,
          std?.weight || "-",
          r.feed_intake_g,
          std?.feed_bird_day || "-",
          r.egg_production,
          std?.egg_production || "-",
          r.egg_weight,
          std?.egg_weight || "-",
          r.feed_company,
          r.feed_type,
          r.remarks,
        ]);
      });
    } else {
      const b = workbook.addWorksheet("Before Sale");
      b.addRow([
        "Visit Date",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mort.",
        "Wt (g)",
        "Std Wt",
        "Feed (Kg)",
        "Feed/B (g)",
        "Std Feed",
        "FCR(w/o)",
        "Std FCR",
        "FCR(With)",
        "Feed",
        "Type",
        "Remarks",
      ]);
      farmerHistory.before.forEach((r: any) => {
        const d = getAgeDays(selectedHistoryVisit.hatch_date, r.visit_date);
        const std = broilerStds[d];
        b.addRow([
          r.visit_date,
          selectedHistoryVisit.hatch_date,
          selectedHistoryVisit.t_birds,
          d,
          r.mortality_pcs,
          r.weight_gm_b,
          std?.weight || "-",
          r.feed_t_kg,
          r.feed_gm_b,
          std?.feed_intake || "-",
          r.fcr_w_out_m,
          std?.fcr || "-",
          r.fcr_with_m,
          r.feed_company,
          r.feed_type,
          r.remarks,
        ]);
      });
    }

    const buf = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buf]),
      `Batch_History_${
        selectedHistoryVisit.batch_label ||
        selectedHistoryVisit.farmers?.farmer_name
      }.xlsx`
    );
  }

  const companyScopedFarmers = statusFarmers.filter((f) => {
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
    { key: "all" as StatusTab, title: "Total", color: "border-slate-300 bg-white text-slate-800", badgeBg: "bg-slate-800 text-white", data: getSummaryForGroup(() => true) },
    { key: "Running" as StatusTab, title: "Running", color: "border-emerald-300 bg-emerald-50/60 text-emerald-900", badgeBg: "bg-emerald-600 text-white", data: getSummaryForGroup((f) => f.computedStatus === "Running") },
    { key: "Active" as StatusTab, title: "Active (<2M)", color: "border-blue-300 bg-blue-50/60 text-blue-900", badgeBg: "bg-blue-600 text-white", data: getSummaryForGroup((f) => f.computedStatus === "Active") },
    { key: "Inactive" as StatusTab, title: "Inactive (>2M)", color: "border-amber-300 bg-amber-50/60 text-amber-900", badgeBg: "bg-amber-600 text-white", data: getSummaryForGroup((f) => f.computedStatus === "Inactive") },
    { key: "Leave from Nourish" as StatusTab, title: "Leave Nourish", color: "border-red-300 bg-red-50/60 text-red-900", badgeBg: "bg-red-600 text-white", data: getSummaryForGroup((f) => f.nourishTag === "Leave from Nourish") },
    { key: "Convert to Nourish" as StatusTab, title: "Convert Nourish", color: "border-teal-300 bg-teal-50/60 text-teal-900", badgeBg: "bg-teal-600 text-white", data: getSummaryForGroup((f) => f.nourishTag === "Convert to Nourish") },
  ];

  const availableAgentsForDropdown =
    selectedCompany === "all"
      ? allAgents
      : allAgents.filter((a) => String(a.company_id) === String(selectedCompany));

  // One row per running batch; one row per non-running farmer
  const tableRows: any[] = [];
  companyScopedFarmers.forEach((f) => {
    if (f.runningVisits.length > 0) {
      f.runningVisits.forEach((v: any) =>
        tableRows.push({
          kind: "batch",
          key: "v" + v.id,
          status: "Running",
          category: v.category || "Broiler",
          farmer: f,
          visit: v,
        })
      );
    } else {
      tableRows.push({
        kind: "farmer",
        key: "f" + f.id,
        status: f.computedStatus,
        category: f.category,
        farmer: f,
        visit: null,
      });
    }
  });

  function rowRank(r: any) {
    const base = STATUS_RANK[r.status] ?? 3;
    // tagged farmers (Leave / Convert) that are not running go after the main statuses
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
          String(r.visit?.batch_label || "").toLowerCase().includes(q);
        if (!match) return false;
      }
      return true;
    })
    .sort(
      (a, b) =>
        rowRank(a) - rowRank(b) ||
        String(a.farmer.farmer_name || "").localeCompare(
          String(b.farmer.farmer_name || "")
        )
    );

  const filteredAgentsList = Object.values(
    tableFilteredRows.reduce((acc: Record<string, any>, r: any) => {
      const key = String(r.farmer.agent_id || "unknown");
      if (!acc[key]) {
        acc[key] = {
          agentName: r.farmer.agentName,
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



  // ---------- Activate / Run Farm / Add Batch ----------
  function toggleFarmerCheckbox(id: string) {
    setSelectedFarmerIds((prev) =>
      prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
    );
  }

  function openActivateModal(farmer: any) {
    const defaultCat = farmer.category || "Broiler";
    const defaultBreeds = BREEDS_BY_CATEGORY[defaultCat] || [];
    const defaultHatchery =
      companies.find((c) => String(c.id) === String(farmer.company_id))
        ?.company_name ||
      companies[0]?.company_name ||
      "";
    setActivateBulk(false);
    setActivateFarmer(farmer);
    setActivateMode("add");
    setActivateForm({
      category: defaultCat,
      hatch_date: new Date().toISOString().split("T")[0],
      hatchery: defaultHatchery,
      breed: defaultBreeds[0] || "",
      box_no: "",
      t_birds: "",
      batch_mark: "",
    });
    setShowActivateModal(true);
  }

  function openBulkActivateModal() {
    if (selectedFarmerIds.length === 0) return;
    const defaultHatchery =
      companies.find((c) => String(c.id) === String(selectedCompany))
        ?.company_name ||
      companies[0]?.company_name ||
      "";
    setActivateBulk(true);
    setActivateFarmer(null);
    setActivateMode("add");
    setActivateForm({
      category: "Broiler",
      hatch_date: new Date().toISOString().split("T")[0],
      hatchery: defaultHatchery,
      breed: BREEDS_BY_CATEGORY["Broiler"][0],
      box_no: "",
      t_birds: "",
      batch_mark: "",
    });
    setShowActivateModal(true);
  }

  const activatePrevCats: string[] = (() => {
    if (activateBulk || !activateFarmer) return [];
    const running = (activateFarmer.runningVisits || [])
      .map((v: any) => v.category)
      .filter(Boolean);
    const uniq = Array.from(new Set(running)) as string[];
    if (uniq.length > 0) return uniq;
    return activateFarmer.category ? [activateFarmer.category] : [];
  })();
  const activateNeedsChoice =
    !activateBulk &&
    activatePrevCats.length > 0 &&
    !activatePrevCats.some(
      (c) => c.toLowerCase() === activateForm.category.toLowerCase()
    );

  async function handleActivateSubmit() {
    if (!activateForm.hatch_date || !activateForm.t_birds) {
      alert("Hatch Date and Total Birds are required.");
      return;
    }
    const targets = activateBulk
      ? statusFarmers.filter((f) => selectedFarmerIds.includes(String(f.id)))
      : activateFarmer
      ? [activateFarmer]
      : [];
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
          !activateBulk && activateNeedsChoice ? activateMode : "add";

        // give unmarked running batches a clear "<Type> Batch (n)" mark
        if (running.length > 0) {
          const counters: Record<string, number> = {};
          for (const rv of running) {
            const rc = rv.category || "Broiler";
            counters[rc] = (counters[rc] || 0) + 1;
            const lbl = String(rv.batch_label || "");
            const unmarked =
              !lbl ||
              lbl === baseName ||
              (lbl.startsWith(baseName) && !lbl.includes("Batch"));
            if (unmarked) {
              await supabase
                .from("vet_visits")
                .update({ batch_label: `${rc} Batch (${counters[rc]})` })
                .eq("id", rv.id);
            }
          }
        }

        const manual = (!activateBulk && activateForm.batch_mark.trim()) || "";
        const batchLabel = manual
          ? manual
          : running.length > 0
          ? `${cat} Batch (${sameCat.length + 1})`
          : baseName;

        // ---- Farmers page details ----
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
          await supabase.from("farmers").update(farmerUpdate).eq("id", target.id);
        }

        const insertPayload: any = {
          farmer_id: String(target.id),
          agent_id: String(target.agent_id),
          company_id: String(target.company_id),
          category: cat,
          hatch_date: activateForm.hatch_date,
          hatchery: activateForm.hatchery,
          breed: activateForm.breed,
          box_no: activateForm.box_no,
          t_birds: qty,
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
      alert("Farm is now Running!");
      setShowActivateModal(false);
      setSelectedFarmerIds([]);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to activate farmer");
    }
  }

  // ---------- Export (status list) ----------
  async function exportFilteredFarmersExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Farmer Status");
    sheet.addRow([
      "SL", "Status", "Date", "Category", "Agent", "Farmer Name", "Mobile", "Address", "Nourish Conversion",
    ]);
    sheet.getRow(1).eachCell((c) => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "0F172A" } };
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    });
    tableFilteredRows.forEach((r, idx) => {
      const f = r.farmer;
      sheet.addRow([
        idx + 1,
        r.status,
        r.kind === "batch" ? r.visit.hatch_date || "-" : f.statusDate || "-",
        r.category,
        f.agentName,
        r.visit?.batch_label || f.farmer_name,
        f.mobile || "-",
        [f.village, f.upazila].filter(Boolean).join(", "),
        f.nourishTag
          ? `${f.nourishTag}${f.nourishTagDate ? ` (${f.nourishTagDate})` : ""}`
          : "-",
      ]);
    });
    sheet.columns.forEach((col) => (col.width = 20));
    const buf = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buf]),
      `Farmer_Status_${selectedStatusFilter}_${new Date().toISOString().slice(0, 10)}.xlsx`
    );
  }

  // ---------- Edit saved report (from History) ----------
  const REPORT_FIELDS: Record<string, { key: string; label: string; type?: string }[]> = {
    before_sale_reports: [
      { key: "visit_date", label: "Visit Date", type: "date" },
      { key: "mortality_pcs", label: "Mortality (pcs)", type: "number" },
      { key: "weight_gm_b", label: "Weight (g/b)", type: "number" },
      { key: "feed_t_kg", label: "Feed Total (Kg)", type: "number" },
      { key: "feed_gm_b", label: "Feed (g/b)", type: "number" },
      { key: "fcr_w_out_m", label: "FCR (w/o M)", type: "number" },
      { key: "fcr_with_m", label: "FCR (with M)", type: "number" },
      { key: "uneven", label: "Uneven" },
      { key: "indigestion", label: "Indigestion" },
      { key: "orange_feces", label: "Orange Feces" },
      { key: "watery_feces", label: "Watery Feces" },
      { key: "feeder", label: "Feeder" },
      { key: "drinker", label: "Drinker" },
      { key: "floor_space", label: "Floor Space" },
      { key: "litter_condition", label: "Litter Condition" },
      { key: "ventilation", label: "Ventilation" },
      { key: "feed_scattering", label: "Feed Scattering" },
      { key: "feed_hardness", label: "Feed Hardness" },
      { key: "feed_company", label: "Feed Company" },
      { key: "feed_type", label: "Feed Type" },
      { key: "pro_date", label: "MFG Date", type: "date" },
      { key: "batch_no", label: "Batch No." },
      { key: "remarks", label: "Remarks" },
    ],
    after_sale_reports: [
      { key: "visit_date", label: "Sale Date", type: "date" },
      { key: "mortality", label: "Mortality", type: "number" },
      { key: "weight", label: "Weight (g)", type: "number" },
      { key: "feed_consumption_kg", label: "Total Feed (Kg)", type: "number" },
      { key: "feed_gm_b", label: "Feed/Bird (g)", type: "number" },
      { key: "feed", label: "Last Feed Type" },
      { key: "feed_mill", label: "Last Feed Company" },
      { key: "fcr_w_out_m", label: "FCR (w/o M)", type: "number" },
      { key: "fcr_with_m", label: "FCR (with M)", type: "number" },
      { key: "remarks", label: "Remarks" },
    ],
    layer_reports: [
      { key: "visit_date", label: "Visit Date", type: "date" },
      { key: "b_weight_g", label: "Body Weight (g)", type: "number" },
      { key: "feed_total_kg", label: "Feed Total (Kg)", type: "number" },
      { key: "feed_intake_g", label: "Feed (g/b/d)", type: "number" },
      { key: "egg_production", label: "Egg Production %", type: "number" },
      { key: "egg_weight", label: "Egg Weight (g)", type: "number" },
      { key: "feed_company", label: "Feed Company" },
      { key: "feed_type", label: "Feed Type" },
      { key: "pro_date", label: "MFG Date", type: "date" },
      { key: "batch_no", label: "Batch No." },
      { key: "hatchery", label: "Hatchery" },
      { key: "breed", label: "Breed" },
      { key: "disease", label: "Disease" },
      { key: "remarks", label: "Remarks" },
    ],
  };

  function openEditReport(table: string, rep: any) {
    const form: Record<string, any> = {};
    (REPORT_FIELDS[table] || []).forEach((fld) => {
      form[fld.key] = rep[fld.key] ?? "";
    });
    setEditReport({ table, id: rep.id, form });
  }

  async function handleSaveEditedReport() {
    if (!editReport) return;
    const fields = REPORT_FIELDS[editReport.table] || [];
    const payload: Record<string, any> = {};
    fields.forEach((fld) => {
      const v = editReport.form[fld.key];
      if (fld.type === "number") payload[fld.key] = Number(v) || 0;
      else if (fld.type === "date") payload[fld.key] = v || null;
      else payload[fld.key] = v ?? "";
    });
    const { error } = await supabase
      .from(editReport.table)
      .update(payload)
      .eq("id", editReport.id);
    if (error) {
      alert(error.message || "Failed to update report");
      return;
    }
    setEditReport(null);
    if (selectedHistoryVisit) await handleHistoryClick(selectedHistoryVisit);
  }

  const activeAgeDays = getAgeDays(
    activeReportVisit?.hatch_date,
    reportVisitDate
  );
  const activeAgeWeeks = getAgeWeeks(
    activeReportVisit?.hatch_date,
    reportVisitDate
  );
  const activeBroilerStd = broilerStds[activeAgeDays] || null;
  const activeLayerStd = getLayerStdForVisit(
    layerForm.breed || activeReportVisit?.breed,
    activeAgeWeeks
  );

  const beforeChangedComp = getFeedChangeCompanyObj(beforeForm.feed_company);
  const afterChangedComp = getFeedChangeCompanyObj(afterForm.feed_mill);
  const layerChangedComp = getFeedChangeCompanyObj(layerForm.feed_company);

  const activeFeedTypeOptions =
    UPDATED_FEED_TYPES_BY_CATEGORY[activeReportVisit?.category || "Broiler"] ||
    UPDATED_FEED_TYPES_BY_CATEGORY["Broiler"];

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-4">
      {/* Top Bar */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-3 py-2 shadow-sm">
        <div className="flex items-center gap-2">
          <h1 className="text-sm sm:text-base font-bold text-slate-800">
            Vet Visits
          </h1>
          <span className="text-[11px] font-semibold bg-slate-100 border border-slate-200 px-2 py-0.5 rounded-lg text-slate-700">
            Rows: <strong>{tableFilteredRows.length}</strong>
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          {selectedFarmerIds.length > 0 && (
            <button
              type="button"
              onClick={openBulkActivateModal}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded-lg text-[11px] font-bold shadow-sm transition"
            >
              Activate Selected ({selectedFarmerIds.length})
            </button>
          )}
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

      {/* Compact status summary + filters */}
      <div className="bg-white p-2 sm:p-2.5 rounded-xl border border-slate-300 shadow-sm space-y-2">
        <div className="flex flex-wrap gap-1.5">
          {summaryCards.map((card) => {
            const isSelected = selectedStatusFilter === card.key;
            return (
              <div
                key={card.key}
                className={`flex items-center gap-1 rounded-lg border pl-2 pr-1 py-0.5 text-[11px] cursor-pointer transition ${card.color} ${
                  isSelected ? "ring-2 ring-emerald-600" : ""
                }`}
                onClick={() => setSelectedStatusFilter(card.key)}
              >
                <span className="font-bold">{card.title}</span>
                <span className={`px-1.5 rounded-full font-extrabold ${card.badgeBg}`}>
                  {card.data.total}
                </span>
                <button
                  type="button"
                  title="Category details"
                  onClick={(e) => {
                    e.stopPropagation();
                    setExpandedCard(expandedCard === card.key ? null : card.key);
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
          <select
            value={selectedCompany}
            onChange={(e) => {
              setSelectedCompany(e.target.value);
              setSelectedAgent("all");
            }}
            className="rounded-md border border-slate-300 bg-white px-1.5 py-1 text-[11px] outline-none"
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
            className="rounded-md border border-slate-300 bg-white px-1.5 py-1 text-[11px] outline-none max-w-[140px]"
          >
            <option value="all">All Agents</option>
            {availableAgentsForDropdown.map((ag) => (
              <option key={ag.id} value={String(ag.id)}>
                {ag.agent_name}
              </option>
            ))}
          </select>
          <select
            value={selectedCategoryFilter}
            onChange={(e) => setSelectedCategoryFilter(e.target.value)}
            className="rounded-md border border-slate-300 bg-white px-1.5 py-1 text-[11px] outline-none"
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
            onChange={(e) => setSelectedStatusFilter(e.target.value as StatusTab)}
            className="rounded-md border border-slate-300 bg-white px-1.5 py-1 text-[11px] outline-none font-semibold"
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
            title="Reset filters"
            onClick={() => {
              setSelectedCompany("all");
              setSelectedAgent("all");
              setSelectedCategoryFilter("all");
              setSelectedStatusFilter("all");
              setSearchQuery("");
              setSelectedFarmerIds([]);
              setFocusVisitId(null);
            }}
            className="rounded-md bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2 py-1 text-[11px] font-semibold text-slate-700 transition"
          >
            Reset
          </button>
        </div>
      </div>

      {/* Agent-wise table: Status | Age | Category first, sorted Running > Active > Inactive > others */}
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
              <div className="bg-slate-100 px-3 py-1.5 border-b border-slate-300 flex flex-wrap justify-between items-center gap-2">
                <h2 className="font-bold text-xs sm:text-sm text-slate-800">
                  Agent:{" "}
                  <span className="text-emerald-700">{agentGroup.agentName}</span>
                </h2>
                <span className="text-[11px] bg-white border border-slate-300 text-slate-700 px-2 py-0.5 rounded font-semibold">
                  {agentGroup.rows.length} ({countsText})
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-50 text-slate-700">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1.5">Status & Date</th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center font-bold text-emerald-700">Age</th>
                      <th className="border border-slate-300 px-2 py-1.5">Category</th>
                      <th className="border border-slate-300 px-2 py-1.5">Farmer</th>
                      <th className="border border-slate-300 px-2 py-1.5">Mobile & Address</th>
                      <th className="border border-slate-300 px-2 py-1.5">H.Date</th>
                      <th className="border border-slate-300 px-2 py-1.5">Hatchery & Breed</th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">Birds</th>
                      <th className="border border-slate-300 px-2 py-1.5">Vaccine / Alert</th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">Take Report</th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">History</th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">Edit</th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {agentGroup.rows.map((row: any) => {
                      const farmer = row.farmer;
                      const visit = row.visit;
                      const isBatch = row.kind === "batch";
                      const category = row.category || "Broiler";
                      const isLayer = category.toLowerCase() === "layer";
                      const ageDays = isBatch ? getAgeDays(visit.hatch_date) : 0;
                      const isSellTime =
                        isBatch && category.toLowerCase() === "broiler" && ageDays >= 28;
                      const vacInfo = isBatch ? getVaccineInfoForVisit(visit) : null;
                      const tag = isBatch ? batchTagOf(visit.batch_label, farmer.farmer_name) : "";
                      const isFocused = isBatch && focusVisitId === String(visit.id);
                      const days = farmer.daysSinceSale;
                      const daysText =
                        days === null || days === undefined
                          ? ""
                          : days === 0
                          ? "today"
                          : `${days} day${days === 1 ? "" : "s"} before`;
                      const afterBtnRing = isFocused
                        ? "ring-4 ring-yellow-400 animate-pulse"
                        : "";

                      return (
                        <tr
                          key={row.key}
                          id={isBatch ? `visit-row-${visit.id}` : undefined}
                          className={`hover:bg-slate-50 ${isFocused ? "bg-yellow-50" : ""}`}
                        >
                          {/* 1. STATUS & DATE */}
                          <td className="border border-slate-300 px-2 py-1.5 whitespace-nowrap">
                            <div className="flex items-center gap-1.5">
                              {!isBatch && (
                                <input
                                  type="checkbox"
                                  checked={selectedFarmerIds.includes(String(farmer.id))}
                                  onChange={() => toggleFarmerCheckbox(String(farmer.id))}
                                  className="w-3.5 h-3.5 accent-emerald-600 cursor-pointer"
                                />
                              )}
                              <span
                                className={`inline-block px-2 py-0.5 rounded-full text-[11px] font-bold ${
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
                                  className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                    farmer.nourishTag === "Convert to Nourish"
                                      ? "bg-teal-100 text-teal-800 border border-teal-300"
                                      : "bg-red-100 text-red-800 border border-red-300"
                                  }`}
                                >
                                  {farmer.nourishTag}
                                </span>
                              )}
                            </div>
                            <div className="text-[10px] text-slate-500 mt-0.5">
                              {isBatch ? (
                                <>H.Date: {formatDmy(visit.hatch_date)}</>
                              ) : farmer.statusDate ? (
                                <>
                                  {farmer.closedType || "Sale/Cull"}: {formatDmy(farmer.statusDate)}
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
                          </td>

                          {/* 2. AGE */}
                          <td className="border border-slate-300 px-2 py-1.5 text-center font-bold text-emerald-700 whitespace-nowrap">
                            {isBatch ? calculateAge(visit.hatch_date, category) : "-"}
                          </td>

                          {/* 3. CATEGORY */}
                          <td className="border border-slate-300 px-2 py-1.5 font-semibold">
                            {category}
                          </td>

                          <td className="border border-slate-300 px-2 py-1.5 font-bold text-slate-800">
                            {farmer.farmer_name || "N/A"}
                            {tag && (
                              <div className="mt-0.5">
                                <span className="px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 text-[10px] font-semibold">
                                  {tag}
                                </span>
                              </div>
                            )}
                          </td>

                          <td className="border border-slate-300 px-2 py-1.5 text-slate-600">
                            <div>{farmer.mobile || "-"}</div>
                            <div className="text-[11px] text-slate-500">
                              {[farmer.village, farmer.upazila].filter(Boolean).join(", ")}
                            </div>
                          </td>

                          {isBatch ? (
                            <>
                              <td className="border border-slate-300 px-2 py-1.5 whitespace-nowrap">
                                {visit.hatch_date || "-"}
                              </td>

                              <td className="border border-slate-300 px-2 py-1.5">
                                <div className="font-medium text-slate-800">{visit.hatchery || "-"}</div>
                                <div className="text-[11px] text-slate-500">{visit.breed || "-"}</div>
                              </td>

                              <td className="border border-slate-300 px-2 py-1.5 text-center font-bold text-slate-800">
                                {visit.t_birds || 0}
                              </td>

                              <td className="border border-slate-300 px-2 py-1.5 min-w-[190px]">
                                <div className="flex items-start justify-between gap-1.5">
                                  <div className="space-y-0.5 flex-1">
                                    {visit.last_vaccine_name ? (
                                      <div className="text-[11px] text-emerald-800 px-1.5 py-0.5 rounded">
                                        <b>Last:</b> {visit.last_vaccine_name}{" "}
                                        {visit.last_vaccine_date ? `(${formatDmy(visit.last_vaccine_date)})` : ""}
                                      </div>
                                    ) : (
                                      <div className="text-[11px] text-slate-400 italic">No vaccine recorded yet</div>
                                    )}
                                    {vacInfo?.upcoming && (
                                      <div className="text-[11px] text-slate-600 bg-emerald-100 border border-emerald-200 mt-0.5">
                                        <b>Next:</b> {vacInfo.upcoming.vaccine_name} (Day {vacInfo.upcoming.age_days} • {formatDmy(vacInfo.upcomingDateStr)})
                                      </div>
                                    )}
                                    {isSellTime && (
                                      <div className="mt-1 inline-block bg-red-100 border border-red-300 text-red-700 text-[11px] px-2 py-0.5 rounded font-bold">
                                        Time to Sale
                                      </div>
                                    )}
                                  </div>
                                  <button
                                    type="button"
                                    onClick={() => handleOpenLogVaccineModal(visit)}
                                    title="Add Given Vaccine"
                                    className="w-6 h-6 shrink-0 rounded-md bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-sm flex items-center justify-center shadow-sm transition"
                                  >
                                    +
                                  </button>
                                </div>
                              </td>

                              <td className="border border-slate-300 px-2 py-1.5 text-center whitespace-nowrap">
                                {isLayer ? (
                                  <button
                                    type="button"
                                    id={`after-sale-btn-${visit.id}`}
                                    onClick={() => openReportModal("layer_report", visit)}
                                    className={`bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold ${afterBtnRing}`}
                                  >
                                    Layer Report
                                  </button>
                                ) : isSellTime ? (
                                  <div className="flex flex-col gap-1">
                                    <button
                                      type="button"
                                      id={`after-sale-btn-${visit.id}`}
                                      onClick={() => openReportModal("after_sale", visit)}
                                      className={`bg-red-600 hover:bg-red-700 text-white px-2.5 py-1 rounded text-[11px] font-bold shadow-sm ${afterBtnRing}`}
                                    >
                                      After Sale Report
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => openReportModal("before_sale", visit)}
                                      className="bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 px-2 py-0.5 rounded text-[10px] font-semibold"
                                    >
                                      Before Sale
                                    </button>
                                  </div>
                                ) : (
                                  <div className="flex flex-col gap-1">
                                    <button
                                      type="button"
                                      onClick={() => openReportModal("before_sale", visit)}
                                      className="bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold shadow-sm"
                                    >
                                      Before Sale Report
                                    </button>
                                    <button
                                      type="button"
                                      id={`after-sale-btn-${visit.id}`}
                                      onClick={() => openReportModal("after_sale", visit)}
                                      className={`bg-slate-100 hover:bg-slate-200 text-slate-600 border border-slate-300 px-2 py-0.5 rounded text-[10px] ${afterBtnRing}`}
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
                              className="border border-slate-300 px-2 py-1.5 text-[11px] text-slate-400 italic"
                            >
                              No running batch
                            </td>
                          )}

                          <td className="border border-slate-300 px-2 py-1.5 text-center">
                            <button
                              type="button"
                              onClick={() => {
                                if (isBatch) {
                                  handleHistoryClick(visit);
                                } else if (farmer.latestClosed) {
                                  handleHistoryClick({
                                    ...farmer.latestClosed,
                                    farmers: farmer,
                                    agents: { agent_name: farmer.agentName },
                                  });
                                } else {
                                  alert("No batch history found for this farmer.");
                                }
                              }}
                              className="bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 px-2 py-1 rounded text-[11px] font-semibold"
                            >
                              History
                            </button>
                          </td>

                          <td className="border border-slate-300 px-2 py-1.5 text-center">
                            {isBatch ? (
                              <button
                                type="button"
                                onClick={() => handleEditClick(visit)}
                                className="bg-amber-50 text-amber-700 border border-amber-200 hover:bg-amber-100 px-2 py-1 rounded text-[11px] font-semibold"
                              >
                                Edit
                              </button>
                            ) : (
                              <span className="text-slate-300">-</span>
                            )}
                          </td>

                          <td className="border border-slate-300 px-2 py-1.5 text-center whitespace-nowrap">
                            <button
                              type="button"
                              onClick={() => openActivateModal(farmer)}
                              className={`px-2.5 py-1 rounded text-[11px] font-semibold shadow-sm text-white ${
                                row.status === "Running"
                                  ? "bg-emerald-600 hover:bg-emerald-700"
                                  : row.status === "Active"
                                  ? "bg-blue-600 hover:bg-blue-700"
                                  : "bg-amber-600 hover:bg-amber-700"
                              }`}
                            >
                              {row.status === "Running"
                                ? "+ Add Batch"
                                : row.status === "Active"
                                ? "Run Farm"
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
          );
        })
      )}

      {/* ==================================================
          ACTIVATE / RUN FARM / ADD BATCH MODAL
      ================================================== */}
      {showActivateModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[60] p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-md max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h3 className="text-sm sm:text-base font-bold text-slate-800">
                {activateBulk
                  ? `Bulk Run Farms (${selectedFarmerIds.length})`
                  : activateFarmer?.runningCount > 0
                  ? `Add New Batch: ${activateFarmer?.farmer_name}`
                  : activateFarmer?.computedStatus === "Active"
                  ? `Run Farm: ${activateFarmer?.farmer_name}`
                  : `Activate Farmer: ${activateFarmer?.farmer_name}`}
              </h3>
              <button onClick={() => setShowActivateModal(false)} className="text-slate-500 font-bold">✖</button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-bold text-slate-600 mb-1">Category / Farm Type *</label>
                <select
                  value={activateForm.category}
                  onChange={(e) => {
                    const cat = e.target.value;
                    const breeds = BREEDS_BY_CATEGORY[cat] || [];
                    setActivateForm({ ...activateForm, category: cat, breed: breeds[0] || "" });
                  }}
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white outline-none"
                >
                  {CATEGORIES.map((cat) => (
                    <option key={cat} value={cat}>{cat}</option>
                  ))}
                </select>
              </div>

              {activateNeedsChoice && (
                <div className="border border-amber-300 bg-amber-50 rounded-xl p-3 space-y-2">
                  <div className="font-bold text-amber-900">
                    Before: {activatePrevCats.join(", ")} → Now: {activateForm.category}
                  </div>
                  <label className="flex items-start gap-2 cursor-pointer">
                    <input type="radio" name="actMode" checked={activateMode === "add"} onChange={() => setActivateMode("add")} className="mt-0.5 accent-amber-600" />
                    <span><b>Add</b> new {activateForm.category} batch (keep {activatePrevCats.join(", ")})</span>
                  </label>
                  <label className="flex items-start gap-2 cursor-pointer">
                    <input type="radio" name="actMode" checked={activateMode === "replace"} onChange={() => setActivateMode("replace")} className="mt-0.5 accent-amber-600" />
                    <span><b>Replace</b> {activatePrevCats.join(", ")} with {activateForm.category} on Farmers page</span>
                  </label>
                </div>
              )}

              <div>
                <label className="block font-bold text-slate-600 mb-1">Hatch Date *</label>
                <input type="date" value={activateForm.hatch_date} onChange={(e) => setActivateForm({ ...activateForm, hatch_date: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none" />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">Hatchery (All Companies) *</label>
                <select value={activateForm.hatchery} onChange={(e) => setActivateForm({ ...activateForm, hatchery: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white outline-none">
                  <option value="">Select Hatchery Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.company_name}>{c.company_name}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">Breed *</label>
                <select value={activateForm.breed} onChange={(e) => setActivateForm({ ...activateForm, breed: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white outline-none">
                  <option value="">Select Breed</option>
                  {(BREEDS_BY_CATEGORY[activateForm.category] || []).map((b) => (
                    <option key={b} value={b}>{b}</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-2.5">
                <div>
                  <label className="block font-bold text-slate-600 mb-1">Box No.</label>
                  <input type="text" placeholder="Box No" value={activateForm.box_no} onChange={(e) => setActivateForm({ ...activateForm, box_no: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none" />
                </div>
                <div>
                  <label className="block font-bold text-slate-600 mb-1">Total Birds / Qty *</label>
                  <input type="number" placeholder="Total Birds" value={activateForm.t_birds} onChange={(e) => setActivateForm({ ...activateForm, t_birds: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none" />
                </div>
              </div>

              {!activateBulk && (
                <div>
                  <label className="block font-bold text-slate-600 mb-1">Batch mark (optional)</label>
                  <input type="text" placeholder="Leave empty for auto e.g. Sonali Batch (1)" value={activateForm.batch_mark} onChange={(e) => setActivateForm({ ...activateForm, batch_mark: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 outline-none" />
                </div>
              )}

              <div className="flex justify-end gap-2 pt-3 border-t">
                <button type="button" onClick={() => setShowActivateModal(false)} className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 font-semibold">Cancel</button>
                <button type="button" onClick={handleActivateSubmit} className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold shadow-sm">
                  Run Farm
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          EDIT SAVED REPORT (from History)
      ================================================== */}
      {editReport && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[75] p-3">
          <div className="bg-white rounded-xl p-4 w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center pb-2 mb-3 border-b">
              <h3 className="text-sm font-bold text-slate-800">Edit Report</h3>
              <button onClick={() => setEditReport(null)} className="text-slate-500 font-bold">✖</button>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-xs">
              {(REPORT_FIELDS[editReport.table] || []).map((fld) => (
                <div key={fld.key} className={fld.key === "remarks" ? "col-span-2 sm:col-span-3" : ""}>
                  <label className="block font-semibold text-slate-600 mb-0.5">{fld.label}</label>
                  <input
                    type={fld.type || "text"}
                    value={editReport.form[fld.key] ?? ""}
                    onChange={(e) =>
                      setEditReport({
                        ...editReport,
                        form: { ...editReport.form, [fld.key]: e.target.value },
                      })
                    }
                    className="w-full rounded-lg border border-slate-300 px-2 py-1.5 outline-none focus:border-emerald-500"
                  />
                </div>
              ))}
            </div>
            <div className="flex justify-end gap-2 pt-3 mt-3 border-t">
              <button type="button" onClick={() => setEditReport(null)} className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold">Cancel</button>
              <button type="button" onClick={handleSaveEditedReport} className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold">Save Changes</button>
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
                  {activeVaccineVisit.batch_label ||
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
                  className="px-4 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white font-bold shadow-sm"
                >
                  OK / Add
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 1: BEFORE SALE REPORT (WITH STD VALUES & WHATSAPP GENERATOR)
      ================================================== */}
      {showBeforeSaleModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white p-4 sm:p-5 rounded-2xl w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-3.5">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-sm sm:text-base font-bold text-slate-800">
                Before Sale Report —{" "}
                {activeReportVisit.batch_label ||
                  activeReportVisit.farmers?.farmer_name}
              </h3>
              <button
                onClick={() => setShowBeforeSaleModal(false)}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-emerald-50/70 border border-emerald-200 rounded-xl p-2.5 text-xs">
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
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="w-full border border-emerald-300 rounded px-1.5 py-0.5 bg-white text-xs font-semibold"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-xs">
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Mortality (Pcs)
                </label>
                <input
                  type="number"
                  value={beforeForm.mortality_pcs}
                  onChange={(e) =>
                    updateBeforeCalculations({ mortality_pcs: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Weight (g/b)
                </label>
                <input
                  type="number"
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
                    activeBroilerStd ? `${activeBroilerStd.weight} g` : "N/A"
                  }
                  className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Total Feed (Kg)
                </label>
                <input
                  type="number"
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
                    activeBroilerStd
                      ? `${activeBroilerStd.feed_intake} g`
                      : "N/A"
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
                  value={activeBroilerStd ? `${activeBroilerStd.fcr}` : "N/A"}
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
                <select
                  value={beforeForm.feed_company}
                  onChange={(e) => {
                    setBeforeForm({
                      ...beforeForm,
                      feed_company: e.target.value,
                    });
                    setTransferAgentId("");
                  }}
                  className="w-full border border-emerald-400 rounded-lg px-2.5 py-1.5 bg-white font-semibold"
                >
                  <option value="">Select Feed Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.company_name}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed Type *
                </label>
                <select
                  value={beforeForm.feed_type}
                  onChange={(e) =>
                    setBeforeForm({ ...beforeForm, feed_type: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white"
                >
                  {activeFeedTypeOptions.map((ft) => (
                    <option key={ft} value={ft}>
                      {ft}
                    </option>
                  ))}
                </select>
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
                          String(a.company_id) === String(beforeChangedComp.id)
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
                  value={beforeForm.remarks}
                  onChange={(e) =>
                    setBeforeForm({ ...beforeForm, remarks: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="flex flex-wrap justify-between items-center gap-2 pt-3 border-t">
              <button
                type="button"
                onClick={() =>
                  generateBroilerWhatsAppText(
                    activeReportVisit,
                    beforeForm,
                    reportVisitDate
                  )
                }
                className="px-3.5 py-1.5 rounded-lg bg-green-600 hover:bg-green-700 text-white text-xs font-bold shadow-sm"
              >
                Generate WhatsApp Msg
              </button>

              <div className="flex gap-2">
                <button
                  onClick={() => setShowBeforeSaleModal(false)}
                  className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold"
                >
                  Cancel
                </button>
                <button
                  onClick={handleBeforeSaleSubmit}
                  className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold"
                >
                  Save Report
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 2: MOBILE-FRIENDLY AFTER SALE REPORT (WITH STD VALUES & SOLD BUTTON)
      ================================================== */}
      {showAfterSaleModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white p-4 sm:p-5 rounded-2xl w-full max-w-lg max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-3.5">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-sm sm:text-base font-bold text-slate-800">
                After Sale Report —{" "}
                {activeReportVisit.batch_label ||
                  activeReportVisit.farmers?.farmer_name}
              </h3>
              <button
                onClick={() => setShowAfterSaleModal(false)}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-red-50/70 border border-red-200 rounded-xl p-2.5 text-xs">
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
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="w-full border border-red-300 rounded px-1.5 py-0.5 bg-white text-xs font-semibold"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2.5 text-xs">
              <div className="col-span-2">
                <label className="block font-bold text-slate-600 mb-1">
                  Total Mortality (Pcs)
                </label>
                <input
                  type="number"
                  value={afterForm.mortality}
                  onChange={(e) =>
                    updateAfterCalculations({ mortality: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Avg Weight (g/bird)
                </label>
                <input
                  type="number"
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
                    activeBroilerStd ? `${activeBroilerStd.weight} g` : "N/A"
                  }
                  className="w-full border border-emerald-200 rounded-lg px-2.5 py-1.5 bg-emerald-50/50 text-emerald-900 font-bold"
                />
              </div>

              <div className="col-span-2">
                <label className="block font-bold text-slate-600 mb-1">
                  Total Feed Consumption (Kg)
                </label>
                <input
                  type="number"
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
                    activeBroilerStd
                      ? `${activeBroilerStd.feed_intake} g`
                      : "N/A"
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
                  value={activeBroilerStd ? `${activeBroilerStd.fcr}` : "N/A"}
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
                  Feed Company
                </label>
                <select
                  value={afterForm.feed_mill}
                  onChange={(e) => {
                    setAfterForm({ ...afterForm, feed_mill: e.target.value });
                    setTransferAgentId("");
                  }}
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white"
                >
                  <option value="">Select Feed Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.company_name}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed Type
                </label>
                <select
                  value={afterForm.feed}
                  onChange={(e) =>
                    setAfterForm({ ...afterForm, feed: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white"
                >
                  {activeFeedTypeOptions.map((ft) => (
                    <option key={ft} value={ft}>
                      {ft}
                    </option>
                  ))}
                </select>
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
                          String(a.company_id) === String(afterChangedComp.id)
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

            <div className="flex justify-end gap-2 pt-3 border-t">
              <button
                onClick={() => setShowAfterSaleModal(false)}
                className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleAfterSaleSoldSubmit}
                className="px-5 py-1.5 rounded-lg bg-red-600 hover:bg-red-700 text-white text-xs font-bold shadow-sm"
              >
                Sold
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 3: LAYER PERFORMANCE REPORT (BREED-WISE STANDARDS, NO VACCINE FIELD)
      ================================================== */}
      {showLayerModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white p-4 sm:p-5 rounded-2xl w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-3.5">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-sm sm:text-base font-bold text-indigo-900">
                Layer Performance Report —{" "}
                {activeReportVisit.batch_label ||
                  activeReportVisit.farmers?.farmer_name}
              </h3>
              <button
                onClick={() => setShowLayerModal(false)}
                className="text-slate-500 font-bold"
              >
                ✖
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-indigo-50/70 border border-indigo-200 rounded-xl p-2.5 text-xs">
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
                  Age (Weeks)
                </span>
                <strong className="text-indigo-700">
                  {activeAgeWeeks} Weeks ({activeAgeDays}d)
                </strong>
              </div>
              <div>
                <span className="text-slate-500 block text-[10px] uppercase font-bold">
                  Visit Date
                </span>
                <input
                  type="date"
                  value={reportVisitDate}
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="w-full border border-indigo-300 rounded px-1.5 py-0.5 bg-white text-xs font-semibold"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-xs">
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Body Weight (g)
                </label>
                <input
                  type="number"
                  value={layerForm.b_weight_g}
                  onChange={(e) =>
                    setLayerForm({ ...layerForm, b_weight_g: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>
              <div className="sm:col-span-2">
                <label className="block font-bold text-indigo-700 mb-1">
                  Std. Weight (Week {activeAgeWeeks} •{" "}
                  {layerForm.breed || activeReportVisit.breed})
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
                  Feed Intake (Total Kg/Day)
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
                  Feed (gm/bird/day)
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
                  Egg Production (%)
                </label>
                <input
                  type="number"
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
                  Egg Weight (g)
                </label>
                <input
                  type="number"
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
                  Feed Company
                </label>
                <select
                  value={layerForm.feed_company}
                  onChange={(e) => {
                    setLayerForm({
                      ...layerForm,
                      feed_company: e.target.value,
                    });
                    setTransferAgentId("");
                  }}
                  className="w-full border border-indigo-300 rounded-lg px-2.5 py-1.5 bg-white font-semibold"
                >
                  <option value="">Select Feed Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.company_name}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed Type
                </label>
                <select
                  value={layerForm.feed_type}
                  onChange={(e) =>
                    setLayerForm({ ...layerForm, feed_type: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white"
                >
                  {UPDATED_FEED_TYPES_BY_CATEGORY["Layer"].map((ft) => (
                    <option key={ft} value={ft}>
                      {ft}
                    </option>
                  ))}
                </select>
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
                          String(a.company_id) === String(layerChangedComp.id)
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
                  value={layerForm.remarks}
                  onChange={(e) =>
                    setLayerForm({ ...layerForm, remarks: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="flex flex-wrap justify-between items-center gap-2 pt-3 border-t">
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() =>
                    generateLayerWhatsAppText(activeReportVisit, layerForm)
                  }
                  className="px-3.5 py-1.5 rounded-lg bg-green-600 hover:bg-green-700 text-white text-xs font-bold shadow-sm"
                >
                  Generate WhatsApp Msg
                </button>

                <button
                  type="button"
                  onClick={() => handleLayerReportSubmit(true)}
                  className="px-3.5 py-1.5 rounded-lg bg-red-600 hover:bg-red-700 text-white text-xs font-bold shadow-sm"
                >
                  Culled
                </button>
              </div>

              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => setShowLayerModal(false)}
                  className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-semibold"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={() => handleLayerReportSubmit(false)}
                  className="px-4 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold"
                >
                  Save Layer Report
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 4: MASTER VACCINE SCHEDULE (ADD/DELETE ONLY FOR SUPER ADMIN)
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
                    className="bg-emerald-600 text-white px-3 rounded font-bold shrink-0"
                  >
                    + Add
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
          MODAL 5: BATCH HISTORY MODAL
      ================================================== */}
      {showHistoryModal && selectedHistoryVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white p-5 rounded-xl w-full max-w-6xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-4">
            <div className="flex flex-wrap justify-between items-center gap-2 border-b pb-2.5">
              <div>
                <h3 className="text-base font-bold text-slate-800">
                  History —{" "}
                  {selectedHistoryVisit.batch_label ||
                    selectedHistoryVisit.farmers?.farmer_name}{" "}
                  ({selectedHistoryVisit.category})
                </h3>
                <p className="text-xs text-slate-500">
                  H.Date: {selectedHistoryVisit.hatch_date || "-"} • Total
                  Birds: {selectedHistoryVisit.t_birds || 0} • Breed:{" "}
                  {selectedHistoryVisit.breed || "-"}
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={exportVisitHistoryExcel}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold"
                >
                  Export Excel
                </button>
                <button
                  onClick={() => setShowHistoryModal(false)}
                  className="px-2.5 py-1 rounded bg-slate-100 text-slate-600 font-bold"
                >
                  ✖
                </button>
              </div>
            </div>

            {farmerHistory.feedLogs.length > 0 && (
              <div className="bg-amber-50 border border-amber-200 rounded-xl p-2.5 text-xs">
                <strong className="text-amber-900">
                  Feed Conversion History:{" "}
                </strong>
                {farmerHistory.feedLogs.map((l: any, idx: number) => (
                  <span key={idx} className="mr-3">
                    [{l.change_date}: {l.from_company} ➔ <b>{l.to_company}</b>]
                  </span>
                ))}
              </div>
            )}

            {(selectedHistoryVisit.category || "").toLowerCase() === "layer" ? (
              <div>
                <h4 className="text-xs font-bold uppercase text-indigo-800 mb-2">
                  Layer Performance History ({farmerHistory.layer.length})
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
                          Egg Wt / Std
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed & Type
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Remarks
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5 text-center">
                          WhatsApp
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {farmerHistory.layer.length === 0 ? (
                        <tr>
                          <td
                            colSpan={11}
                            className="border border-slate-300 py-4 text-center text-slate-400"
                          >
                            No layer reports found for this batch.
                          </td>
                        </tr>
                      ) : (
                        farmerHistory.layer.map((rep: any, i: number) => {
                          const wks = getAgeWeeks(
                            selectedHistoryVisit.hatch_date,
                            rep.visit_date
                          );
                          const std = getLayerStdForVisit(
                            selectedHistoryVisit.breed,
                            wks
                          );
                          return (
                            <tr key={i} className="hover:bg-slate-50">
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.visit_date || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {selectedHistoryVisit.hatch_date || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {selectedHistoryVisit.t_birds || 0}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                {wks}w
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b>{rep.b_weight_g}</b> / {std?.weight || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b>{rep.feed_intake_g}</b> /{" "}
                                {std?.feed_bird_day || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b>{rep.egg_production}%</b> /{" "}
                                {std?.egg_production || "-"}%
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <b>{rep.egg_weight}</b> /{" "}
                                {std?.egg_weight || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.feed_company || "-"}{" "}
                                {rep.feed_type ? `(${rep.feed_type})` : ""}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.remarks || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                <button
                                  type="button"
                                  onClick={() =>
                                    generateLayerWhatsAppText(
                                      selectedHistoryVisit,
                                      rep
                                    )
                                  }
                                  className="bg-green-600 text-white px-2 py-0.5 rounded text-[10px] font-bold"
                                >
                                  WtsApp
                                </button>
                                <button
                                  type="button"
                                  onClick={() => openEditReport("layer_reports", rep)}
                                  className="ml-1 bg-amber-500 text-white px-2 py-0.5 rounded text-[10px] font-bold"
                                >
                                  Edit
                                </button>
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
                  <h4 className="text-xs font-bold uppercase text-emerald-800 mb-2">
                    Before Sale Reports ({farmerHistory.before.length})
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
                            FCR(With)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed & Type
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5 text-center">
                            WhatsApp
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {farmerHistory.before.length === 0 ? (
                          <tr>
                            <td
                              colSpan={12}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No before sale reports found.
                            </td>
                          </tr>
                        ) : (
                          farmerHistory.before.map((rep: any, i: number) => {
                            const d = getAgeDays(
                              selectedHistoryVisit.hatch_date,
                              rep.visit_date
                            );
                            const std = broilerStds[d];
                            return (
                              <tr key={i} className="hover:bg-slate-50">
                                <td className="border border-slate-300 px-2 py-1">
                                  {rep.visit_date}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {selectedHistoryVisit.hatch_date}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {selectedHistoryVisit.t_birds}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                  {d}d
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {rep.mortality_pcs}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{rep.weight_gm_b}</b> /{" "}
                                  {std?.weight || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {rep.feed_t_kg}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{rep.feed_gm_b}</b> /{" "}
                                  {std?.feed_intake || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                  {rep.fcr_w_out_m} / {std?.fcr || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {rep.fcr_with_m}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {rep.feed_company || "-"}{" "}
                                  {rep.feed_type ? `(${rep.feed_type})` : ""}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <button
                                    type="button"
                                    onClick={() =>
                                      generateBroilerWhatsAppText(
                                        selectedHistoryVisit,
                                        rep,
                                        rep.visit_date
                                      )
                                    }
                                    className="bg-green-600 text-white px-2 py-0.5 rounded text-[10px] font-bold"
                                  >
                                    WtsApp
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => openEditReport("before_sale_reports", rep)}
                                    className="ml-1 bg-amber-500 text-white px-2 py-0.5 rounded text-[10px] font-bold"
                                  >
                                    Edit
                                  </button>
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
                    After Sale Reports ({farmerHistory.after.length})
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
                            Feed (Kg)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed/B / Std (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR(w/o) / Std
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR(With)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Last Feed Mill & Type
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5 text-center">
                            Edit
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {farmerHistory.after.length === 0 ? (
                          <tr>
                            <td
                              colSpan={12}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No after sale reports found.
                            </td>
                          </tr>
                        ) : (
                          farmerHistory.after.map((rep: any, i: number) => {
                            const d = getAgeDays(
                              selectedHistoryVisit.hatch_date,
                              rep.visit_date
                            );
                            const std = broilerStds[d];
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
                              <tr key={i} className="hover:bg-slate-50">
                                <td className="border border-slate-300 px-2 py-1">
                                  {rep.visit_date}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {selectedHistoryVisit.hatch_date}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {tB}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                  {d}d
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {rep.mortality}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{rep.weight}</b> / {std?.weight || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {rep.feed_consumption_kg}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <b>{fGm}</b> / {std?.feed_intake || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                  {rep.fcr_w_out_m} / {std?.fcr || "-"}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  {rep.fcr_with_m}
                                </td>
                                <td className="border border-slate-300 px-2 py-1">
                                  {rep.feed_mill || "-"}{" "}
                                  {rep.feed ? `(${rep.feed})` : ""}
                                </td>
                                <td className="border border-slate-300 px-2 py-1 text-center">
                                  <button
                                    type="button"
                                    onClick={() => openEditReport("after_sale_reports", rep)}
                                    className="bg-amber-500 text-white px-2 py-0.5 rounded text-[10px] font-bold"
                                  >
                                    Edit
                                  </button>
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
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[70] p-3">
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

      {/* ==================================================
          MODAL 7: EDIT VISIT MODAL
      ================================================== */}
      {showEditModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white p-5 rounded-xl w-full max-w-md shadow-2xl border border-slate-300 space-y-3 text-xs">
            <div className="flex justify-between items-center border-b pb-2">
              <h3 className="text-sm font-bold text-slate-800">
                Edit Batch Details
              </h3>
              <button onClick={() => setShowEditModal(false)}>✖</button>
            </div>

            <div>
              <label className="block font-bold text-slate-600 mb-1">
                Category
              </label>
              <select
                value={editForm.category}
                onChange={(e) => {
                  const cat = e.target.value;
                  setEditForm({
                    ...editForm,
                    category: cat,
                    breed: (BREEDS_BY_CATEGORY[cat] || [])[0] || "",
                  });
                }}
                className="w-full rounded border px-2.5 py-1.5 bg-white"
              >
                {CATEGORIES.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block font-bold text-slate-600 mb-1">
                Hatch Date
              </label>
              <input
                type="date"
                value={editForm.hatch_date}
                onChange={(e) =>
                  setEditForm({ ...editForm, hatch_date: e.target.value })
                }
                className="w-full rounded border px-2.5 py-1.5"
              />
            </div>

            <div>
              <label className="block font-bold text-slate-600 mb-1">
                Hatchery
              </label>
              <select
                value={editForm.hatchery}
                onChange={(e) =>
                  setEditForm({ ...editForm, hatchery: e.target.value })
                }
                className="w-full rounded border px-2.5 py-1.5 bg-white"
              >
                <option value="">Select Hatchery</option>
                {companies.map((c) => (
                  <option key={c.id} value={c.company_name}>
                    {c.company_name}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block font-bold text-slate-600 mb-1">
                Breed
              </label>
              <select
                value={editForm.breed}
                onChange={(e) =>
                  setEditForm({ ...editForm, breed: e.target.value })
                }
                className="w-full rounded border px-2.5 py-1.5 bg-white"
              >
                <option value="">Select Breed</option>
                {(BREEDS_BY_CATEGORY[editForm.category] || []).map((b) => (
                  <option key={b} value={b}>
                    {b}
                  </option>
                ))}
              </select>
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Box No
                </label>
                <input
                  type="text"
                  value={editForm.box_no}
                  onChange={(e) =>
                    setEditForm({ ...editForm, box_no: e.target.value })
                  }
                  className="w-full rounded border px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Total Birds
                </label>
                <input
                  type="number"
                  value={editForm.t_birds}
                  onChange={(e) =>
                    setEditForm({ ...editForm, t_birds: e.target.value })
                  }
                  className="w-full rounded border px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-3 border-t">
              <button
                onClick={() => setShowEditModal(false)}
                className="px-4 py-1.5 rounded bg-slate-200 text-slate-700 font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleUpdateSubmit}
                className="px-4 py-1.5 rounded bg-emerald-600 text-white font-semibold"
              >
                Update
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}