import React, { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import { BREEDS_BY_CATEGORY, CATEGORIES } from "../farmers/FarmerStatusPage";
import { useAuth } from "../../context/AuthContext";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

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

  const [visitsData, setVisitsData] = useState<any[]>([]);
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
  }, [selectedCompany]);

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
      let query = supabase
        .from("vet_visits")
        .select("*")
        .in("status", ["Activated", "Running"]);

      if (selectedCompany && selectedCompany !== "all") {
        query = query.eq("company_id", String(selectedCompany));
      }

      const { data: visits, error: visitErr } = await query;
      if (visitErr) throw visitErr;

      if (!visits || visits.length === 0) {
        setVisitsData([]);
        setLoading(false);
        return;
      }

      const [{ data: agentsData }, { data: farmersData }] = await Promise.all([
        supabase.from("agents").select("*"),
        supabase.from("farmers").select("*"),
      ]);

      const enrichedVisits = visits.map((v) => {
        const agent = agentsData?.find(
          (a) => String(a.id) === String(v.agent_id)
        );
        const farmer = farmersData?.find(
          (f) => String(f.id) === String(v.farmer_id)
        );

        return {
          ...v,
          agents: agent || { agent_name: "Unknown Agent" },
          farmers: farmer || {
            farmer_name: "Unknown Farmer",
            mobile: "-",
            village: "-",
            upazila: "-",
          },
        };
      });

      setVisitsData(enrichedVisits);
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

  const totalRunningCount = visitsData.length;
  const categoryCounts: Record<string, number> = {};
  CATEGORIES.forEach((cat) => {
    categoryCounts[cat] = visitsData.filter(
      (v) => (v.category || "Broiler").toLowerCase() === cat.toLowerCase()
    ).length;
  });

  const agentMap: Record<
    string,
    { agentName: string; visits: any[]; counts: Record<string, number> }
  > = {};

  visitsData.forEach((v) => {
    const agentId = v.agents?.id || "unknown";
    const agentName = v.agents?.agent_name || "Unknown Agent";

    if (!agentMap[agentId]) {
      agentMap[agentId] = { agentName, visits: [], counts: {} };
    }

    if (
      selectedCategoryFilter === "all" ||
      (v.category || "Broiler").toLowerCase() ===
        selectedCategoryFilter.toLowerCase()
    ) {
      agentMap[agentId].visits.push(v);
      const cat = v.category || "Broiler";
      agentMap[agentId].counts[cat] = (agentMap[agentId].counts[cat] || 0) + 1;
    }
  });

  const filteredAgentsList = Object.values(agentMap).filter(
    (ag) => ag.visits.length > 0
  );

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
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-4 py-2.5 shadow-sm">
        <h1 className="text-base sm:text-lg font-bold text-slate-800">
          Vet Visits — Running Farmers
        </h1>

        <button
          type="button"
          onClick={() => setShowVaccineModal(true)}
          className="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
        >
          Vaccine Schedule
        </button>
      </div>

      {/* Category-wise Running Farmers Summary Bar */}
      <div className="bg-white p-3.5 rounded-xl border border-slate-300 shadow-sm space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2 bg-emerald-50/70 border border-emerald-200 rounded-lg px-3.5 py-2">
          <div className="text-xs sm:text-sm font-bold text-emerald-950">
            Total Running:{" "}
            <span className="bg-emerald-600 text-white px-2.5 py-0.5 rounded-full text-xs ml-1">
              {totalRunningCount}
            </span>
          </div>

          <div className="flex flex-wrap gap-1.5 text-xs">
            {CATEGORIES.map((cat) => (
              <button
                key={cat}
                type="button"
                onClick={() =>
                  setSelectedCategoryFilter(
                    selectedCategoryFilter === cat ? "all" : cat
                  )
                }
                className={`px-2.5 py-1 rounded-md border font-semibold transition ${
                  selectedCategoryFilter === cat
                    ? "bg-emerald-600 text-white border-emerald-600"
                    : "bg-white text-slate-700 border-emerald-200 hover:bg-emerald-100"
                }`}
              >
                {cat}: <strong>{categoryCounts[cat] || 0}</strong>
              </button>
            ))}
          </div>
        </div>

        <div className="flex flex-wrap gap-3 items-center">
          <div className="flex items-center gap-2">
            <label className="text-xs font-bold text-slate-600">Company:</label>
            <select
              value={selectedCompany}
              onChange={(e) => setSelectedCompany(e.target.value)}
              className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs outline-none bg-white"
            >
              <option value="all">All Companies</option>
              {companies.map((comp) => (
                <option key={comp.id} value={String(comp.id)}>
                  {comp.company_name}
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-2">
            <label className="text-xs font-bold text-slate-600">
              Category:
            </label>
            <select
              value={selectedCategoryFilter}
              onChange={(e) => setSelectedCategoryFilter(e.target.value)}
              className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs outline-none bg-white"
            >
              <option value="all">All Categories</option>
              {CATEGORIES.map((cat) => (
                <option key={cat} value={cat}>
                  {cat}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* Agent-wise Minimal Bordered Table */}
      {filteredAgentsList.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-400">
          No running farmers found for the selected filter.
        </div>
      ) : (
        filteredAgentsList.map((agentGroup, idx) => {
          const agentCountsText = Object.entries(agentGroup.counts)
            .map(([c, cnt]) => `${c}: ${cnt}`)
            .join(", ");

          return (
            <div
              key={idx}
              className="bg-white border border-slate-300 rounded-xl shadow-sm overflow-hidden"
            >
              <div className="bg-slate-100 px-4 py-2 border-b border-slate-300 flex flex-wrap justify-between items-center gap-2">
                <h2 className="font-bold text-xs sm:text-sm text-slate-800">
                  Agent:{" "}
                  <span className="text-emerald-700">
                    {agentGroup.agentName}
                  </span>
                </h2>
                <span className="text-xs bg-white border border-slate-300 text-slate-700 px-2.5 py-0.5 rounded font-semibold">
                  Running: {agentGroup.visits.length} ({agentCountsText})
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-50 text-slate-700">
                    <tr>
                      <th className="border border-slate-300 px-2.5 py-1.5">
                        Farmer
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5">
                        Mobile & Address
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5">
                        Category
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5">
                        H.Date
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5">
                        Hatchery & Breed
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                        Birds
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center font-bold text-emerald-700">
                        Age
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5">
                        Vaccine / Alert
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                        Take Report
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                        History
                      </th>
                      <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                        Edit
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {agentGroup.visits.map((visit) => {
                      const farmer = visit.farmers || {};
                      const category = visit.category || "Broiler";
                      const isLayer = category.toLowerCase() === "layer";
                      const ageText = calculateAge(visit.hatch_date, category);
                      const ageDays = getAgeDays(visit.hatch_date);
                      const isSellTime =
                        category.toLowerCase() === "broiler" && ageDays >= 28;
                      const vacInfo = getVaccineInfoForVisit(visit);

                      const displayName =
                        visit.batch_label || farmer.farmer_name || "N/A";

                      return (
                        <tr key={visit.id} className="hover:bg-slate-50">
                          <td className="border border-slate-300 px-2.5 py-1.5 font-bold text-slate-800">
                            {displayName}
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 text-slate-600">
                            <div>{farmer.mobile || "-"}</div>
                            <div className="text-[11px] text-slate-500">
                              {[farmer.village, farmer.upazila]
                                .filter(Boolean)
                                .join(", ")}
                            </div>
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 font-semibold">
                            {category}
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 whitespace-nowrap">
                            {visit.hatch_date || "-"}
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5">
                            <div className="font-medium text-slate-800">
                              {visit.hatchery || "-"}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {visit.breed || "-"}
                            </div>
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 text-center font-bold text-slate-800">
                            {visit.t_birds || 0}
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 text-center font-bold text-emerald-700 whitespace-nowrap">
                            {ageText}
                          </td>

                          {/* Vaccine / Alert Column with "+" Button to Add Given Vaccine */}
                         {/* 2. Table Vaccine Column UI (Only Last Given & Next Vaccine) */}
                            <td className="border border-slate-300 px-2.5 py-1.5 min-w-[200px]">
                              <div className="flex items-start justify-between gap-1.5">
                                <div className="space-y-0.5 flex-1">
                                  {/* Last Given Vaccine */}
                                  {visit.last_vaccine_name ? (
                                    <div className="text-[11px] text-emerald-800  px-1.5 py-0.5 rounded">
                                      <b>Last:</b> {visit.last_vaccine_name}{" "}
                                      {visit.last_vaccine_date
                                        ? `(${formatDmy(visit.last_vaccine_date)})`
                                        : ""}
                                    </div>
                                  ) : (
                                    <div className="text-[11px] text-slate-400 italic">
                                      No vaccine recorded yet
                                    </div>
                                  )}

                                  {/* Next Scheduled Vaccine */}
                                  {vacInfo.upcoming && (
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

                                {/* "+" Button to Record Vaccine Given */}
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

                          <td className="border border-slate-300 px-2.5 py-1.5 text-center whitespace-nowrap">
                            {isLayer ? (
                              <button
                                type="button"
                                onClick={() =>
                                  openReportModal("layer_report", visit)
                                }
                                className="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold"
                              >
                                Layer Report
                              </button>
                            ) : isSellTime ? (
                              <div className="flex flex-col gap-1">
                                <button
                                  type="button"
                                  onClick={() =>
                                    openReportModal("after_sale", visit)
                                  }
                                  className="bg-red-600 hover:bg-red-700 text-white px-2.5 py-1 rounded text-[11px] font-bold shadow-sm"
                                >
                                  After Sale Report
                                </button>
                                <button
                                  type="button"
                                  onClick={() =>
                                    openReportModal("before_sale", visit)
                                  }
                                  className="bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 px-2 py-0.5 rounded text-[10px] font-semibold"
                                >
                                  Before Sale
                                </button>
                              </div>
                            ) : (
                              <div className="flex flex-col gap-1">
                                <button
                                  type="button"
                                  onClick={() =>
                                    openReportModal("before_sale", visit)
                                  }
                                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold shadow-sm"
                                >
                                  Before Sale Report
                                </button>
                                <button
                                  type="button"
                                  onClick={() =>
                                    openReportModal("after_sale", visit)
                                  }
                                  className="bg-slate-100 hover:bg-slate-200 text-slate-600 border border-slate-300 px-2 py-0.5 rounded text-[10px]"
                                >
                                  After Sale
                                </button>
                              </div>
                            )}
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 text-center">
                            <button
                              type="button"
                              onClick={() => handleHistoryClick(visit)}
                              className="bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 px-2.5 py-1 rounded text-[11px] font-semibold"
                            >
                              History
                            </button>
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 text-center">
                            <button
                              type="button"
                              onClick={() => handleEditClick(visit)}
                              className="bg-amber-50 text-amber-700 border border-amber-200 hover:bg-amber-100 px-2.5 py-1 rounded text-[11px] font-semibold"
                            >
                              Edit
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
                            Feed Mill & Type
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {farmerHistory.after.length === 0 ? (
                          <tr>
                            <td
                              colSpan={11}
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