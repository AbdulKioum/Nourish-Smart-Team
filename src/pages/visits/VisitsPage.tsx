import React, { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import { BREEDS_BY_CATEGORY, CATEGORIES } from "../farmers/FarmerStatusPage";
import { Syringe, History, Edit3, Plus, Trash2, X } from "lucide-react";

interface VaccineItem {
  id: number;
  category: string;
  age_days: number;
  vaccine_name: string;
  vaccine_type: string;
  route: string;
}

export default function VetVisitsPage() {
  const [companies, setCompanies] = useState<any[]>([]);
  const [allAgents, setAllAgents] = useState<any[]>([]);
  const [selectedCompany, setSelectedCompany] = useState<string>("all");
  const [selectedCategoryFilter, setSelectedCategoryFilter] =
    useState<string>("all");

  const [visitsData, setVisitsData] = useState<any[]>([]);
  const [vaccines, setVaccines] = useState<VaccineItem[]>([]);
  const [_loading, setLoading] = useState(true);

  // Vaccine Schedule Modal State
  const [showVaccineModal, setShowVaccineModal] = useState(false);
  const [vaccineForm, setVaccineForm] = useState({
    category: "Layer",
    age_days: "",
    vaccine_name: "",
    vaccine_type: "Live",
    route: "Eye Drop",
  });

  // Edit Modal State
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

  // History Modal State
  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [selectedHistoryVisit, setSelectedHistoryVisit] = useState<any>(null);
  const [farmerHistory, setFarmerHistory] = useState<{
    before: any[];
    after: any[];
    layer: any[];
  }>({ before: [], after: [], layer: [] });

  // Report Modal States
  const [showBeforeSaleModal, setShowBeforeSaleModal] = useState(false);
  const [showAfterSaleModal, setShowAfterSaleModal] = useState(false);
  const [showLayerModal, setShowLayerModal] = useState(false);
  const [activeReportVisit, setActiveReportVisit] = useState<any>(null);

  const [reportVisitDate, setReportVisitDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );

  // Feed Change -> Agent Transfer State inside Report Modals
  const [transferAgentId, setTransferAgentId] = useState<string>("");

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
    c_val: "",
    e_val: "",
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
    feed: "",
    feed_mill: "",
    fcr_w_out_m: "",
    fcr_with_m: "",
    remarks: "",
  });

  // Layer Form
  const [layerForm, setLayerForm] = useState({
    b_weight_g: "",
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
    vaccine: "",
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
    const [{ data: compData }, { data: agData }, { data: vacData }] =
      await Promise.all([
        supabase.from("companies").select("*").order("company_name"),
        supabase.from("agents").select("*").order("agent_name"),
        supabase
          .from("vaccines")
          .select("*")
          .order("age_days", { ascending: true }),
      ]);

    const comps = compData || [];
    setCompanies(comps);
    setAllAgents(agData || []);
    setVaccines((vacData as VaccineItem[]) || []);

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
    setVaccines((data as VaccineItem[]) || []);
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

  function getAgeDays(hatchDateStr: string, refDateStr?: string): number {
    if (!hatchDateStr) return 0;
    const hatchDate = new Date(hatchDateStr);
    const refDate = refDateStr ? new Date(refDateStr) : new Date();
    const diffTime = refDate.getTime() - hatchDate.getTime();
    return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
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

  // Get Current & Upcoming Vaccine for a Batch based on Category & Age
  function getVaccineInfoForVisit(visit: any) {
    const cat = visit.category || "Broiler";
    const ageDays = getAgeDays(visit.hatch_date);
    const catVaccines = vaccines
      .filter((v) => v.category?.toLowerCase() === cat.toLowerCase())
      .sort((a, b) => a.age_days - b.age_days);

    if (catVaccines.length === 0) return { dueToday: null, upcoming: null };

    const dueToday = catVaccines.find(
      (v) => v.age_days >= ageDays - 1 && v.age_days <= ageDays + 1
    );
    const upcoming = catVaccines.find((v) => v.age_days > ageDays + 1);

    let upcomingDateStr = "";
    if (upcoming && visit.hatch_date) {
      const d = new Date(visit.hatch_date);
      d.setDate(d.getDate() + upcoming.age_days);
      upcomingDateStr = d.toISOString().slice(0, 10);
    }

    return { dueToday, upcoming, upcomingDateStr };
  }

  // Auto Calculations for Before Sale
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
    let fcrWOutM = beforeForm.fcr_w_out_m;
    let fcrWithM = beforeForm.fcr_with_m;

    if (tBirds > 0) {
      feedGmB = ((feedTKg * 1000) / tBirds).toFixed(2);
    }

    if (weightGmB > 0 && feedTKg > 0) {
      const liveBirds = tBirds - mortality;
      const totalWeightKgWithoutM = (liveBirds * weightGmB) / 1000;
      const totalWeightKgWithM = (tBirds * weightGmB) / 1000;

      if (totalWeightKgWithoutM > 0) {
        fcrWOutM = (feedTKg / totalWeightKgWithoutM).toFixed(3);
      }
      if (totalWeightKgWithM > 0) {
        fcrWithM = (feedTKg / totalWeightKgWithM).toFixed(3);
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

  // Auto Calculations for After Sale
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

    let fcrWOutM = afterForm.fcr_w_out_m;
    let fcrWithM = afterForm.fcr_with_m;

    if (weightGm > 0 && feedKg > 0) {
      const liveBirds = tBirds - mortality;
      const totalWeightKgWithoutM = (liveBirds * weightGm) / 1000;
      const totalWeightKgWithM = (tBirds * weightGm) / 1000;

      if (totalWeightKgWithoutM > 0) {
        fcrWOutM = (feedKg / totalWeightKgWithoutM).toFixed(3);
      }
      if (totalWeightKgWithM > 0) {
        fcrWithM = (feedKg / totalWeightKgWithM).toFixed(3);
      }
    }

    setAfterForm((prev) => ({
      ...prev,
      ...updatedFields,
      fcr_w_out_m: fcrWOutM,
      fcr_with_m: fcrWithM,
    }));
  }

  // Check if selected Feed Company in Report is different from Farmer's current Company
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

  // Handle Feed Company Transfer + Leave/Convert Nourish Tracking
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

    if (wasNourish && !isNowNourish) {
      farmerUpdates.nourish_status = "Leave from Nourish";
      farmerUpdates.nourish_status_date = visitDateStr;
    } else if (!wasNourish && isNowNourish) {
      farmerUpdates.nourish_status = "Convert to Nourish";
      farmerUpdates.nourish_status_date = visitDateStr;
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
        c_val: "",
        e_val: "",
        litter_condition: "",
        ventilation: "",
        feed_scattering: "",
        feed_hardness: "",
        feed_company: currentCompName,
        feed_type: "",
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
        feed: "",
        feed_mill: currentCompName,
        fcr_w_out_m: "",
        fcr_with_m: "",
        remarks: "",
      });
      setShowAfterSaleModal(true);
    } else if (type === "layer_report") {
      setLayerForm({
        b_weight_g: "",
        feed_intake_g: "",
        egg_production: "",
        egg_weight: "",
        feed_company: currentCompName,
        feed_type: "",
        pro_date: "",
        batch_no: "",
        hatchery: visit.hatchery || "",
        breed: visit.breed || "",
        disease: "",
        vaccine: "",
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
          c_val: beforeForm.c_val,
          e_val: beforeForm.e_val,
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
      alert("Before Sale Report saved successfully!");
      setShowBeforeSaleModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save report");
    }
  }

  // Mark as SOLD in After Sale Report -> Remove from Running Vet Visits & Set Farmer Active (<2M)
  async function handleAfterSaleSoldSubmit() {
    if (!activeReportVisit) return;
    try {
      await processFeedCompanyChangeIfNeeded(
        activeReportVisit,
        afterForm.feed_mill,
        transferAgentId,
        reportVisitDate
      );

      const { error } = await supabase.from("after_sale_reports").insert([
        {
          vet_visit_id: activeReportVisit.id,
          farmer_id: String(activeReportVisit.farmer_id),
          agent_id: String(transferAgentId || activeReportVisit.agent_id),
          company_id: String(activeReportVisit.company_id),
          visit_date: reportVisitDate,
          mortality: Number(afterForm.mortality) || 0,
          weight: Number(afterForm.weight) || 0,
          feed_consumption_kg: Number(afterForm.feed_consumption_kg) || 0,
          feed: afterForm.feed,
          feed_mill: afterForm.feed_mill,
          fcr_w_out_m: Number(afterForm.fcr_w_out_m) || 0,
          fcr_with_m: Number(afterForm.fcr_with_m) || 0,
          remarks: afterForm.remarks,
        },
      ]);

      if (error) throw error;

      // Mark this visit batch as Sold so it leaves Vet Visits page
      await supabase
        .from("vet_visits")
        .update({ status: "Sold", closed_date: reportVisitDate })
        .eq("id", activeReportVisit.id);

      // Update farmer's last_sale_date & status to Active
      await supabase
        .from("farmers")
        .update({
          status: "Active",
          last_sale_date: reportVisitDate,
        })
        .eq("id", activeReportVisit.farmer_id);

      alert(
        "Batch marked as SOLD! Report saved and farmer moved to Active status."
      );
      setShowAfterSaleModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save after sale report");
    }
  }

  // Layer Report Save OR Mark as Culled
  async function handleLayerReportSubmit(markAsCulled = false) {
    if (!activeReportVisit) return;
    try {
      await processFeedCompanyChangeIfNeeded(
        activeReportVisit,
        layerForm.feed_company,
        transferAgentId,
        reportVisitDate
      );

      const { error } = await supabase.from("layer_reports").insert([
        {
          vet_visit_id: activeReportVisit.id,
          farmer_id: String(activeReportVisit.farmer_id),
          agent_id: String(transferAgentId || activeReportVisit.agent_id),
          company_id: String(activeReportVisit.company_id),
          visit_date: reportVisitDate,
          b_weight_g: Number(layerForm.b_weight_g) || 0,
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
          vaccine: layerForm.vaccine,
          remarks: layerForm.remarks,
        },
      ]);

      if (error) throw error;

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

        alert(
          "Layer Batch marked as CULLED! Report saved and farmer moved to Active status."
        );
      } else {
        alert("Layer Performance Report saved successfully!");
      }

      setShowLayerModal(false);
      loadVetVisits();
    } catch (err: any) {
      alert(err.message || "Failed to save layer report");
    }
  }

  async function handleAddVaccine(e: React.FormEvent) {
    e.preventDefault();
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
      alert(error.message);
      return;
    }
    setVaccineForm({ ...vaccineForm, age_days: "", vaccine_name: "" });
    loadVaccinesList();
  }

  async function handleDeleteVaccine(id: number) {
    await supabase.from("vaccines").delete().eq("id", id);
    loadVaccinesList();
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

    if (isLayer) {
      // Strictly load ONLY Layer Reports for this specific Layer batch
      const { data: lData } = await supabase
        .from("layer_reports")
        .select("*")
        .eq("vet_visit_id", String(visit.id))
        .order("visit_date", { ascending: false });

      setFarmerHistory({
        before: [],
        after: [],
        layer: lData || [],
      });
    } else {
      // Strictly load ONLY Before & After Sale Reports for this specific Broiler/Sonali batch
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
      });
    }
  }

  // Category-wise Running Farmer Counts
  const totalRunningCount = visitsData.length;
  const categoryCounts: Record<string, number> = {};
  CATEGORIES.forEach((cat) => {
    categoryCounts[cat] = visitsData.filter(
      (v) => (v.category || "Broiler").toLowerCase() === cat.toLowerCase()
    ).length;
  });

  // Group by Agent
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

  const beforeChangedComp = getFeedChangeCompanyObj(beforeForm.feed_company);
  const afterChangedComp = getFeedChangeCompanyObj(afterForm.feed_mill);
  const layerChangedComp = getFeedChangeCompanyObj(layerForm.feed_company);

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
          className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
        >
          <Syringe size={14} /> Vaccine Schedule (Broiler / Layer / Sonali)
        </button>
      </div>

      {/* Category-wise Running Farmers Summary Bar + Filters */}
      <div className="bg-white p-3.5 rounded-xl border border-slate-300 shadow-sm space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2 bg-emerald-50/70 border border-emerald-200 rounded-lg px-3.5 py-2">
          <div className="text-xs sm:text-sm font-bold text-emerald-950">
            Total Running Batches:{" "}
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
                        Vaccine / Sale Alert
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

                          <td className="border border-slate-300 px-2.5 py-1.5">
                            {vacInfo.dueToday && (
                              <div className="mb-1 px-2 py-0.5 rounded bg-amber-100 border border-amber-300 text-amber-900 text-[11px] font-bold">
                                💉 Due Now (Day {vacInfo.dueToday.age_days}):{" "}
                                {vacInfo.dueToday.vaccine_name} (
                                {vacInfo.dueToday.route})
                              </div>
                            )}
                            {vacInfo.upcoming && (
                              <div className="text-[11px] text-slate-600">
                                Next: <b>{vacInfo.upcoming.vaccine_name}</b> on{" "}
                                Day {vacInfo.upcoming.age_days} (
                                {vacInfo.upcomingDateStr})
                              </div>
                            )}
                            {isSellTime && (
                              <div className="mt-1 inline-block bg-red-100 border border-red-300 text-red-700 text-[11px] px-2 py-0.5 rounded font-bold">
                                🔔 Time to Sale (28+ Days)
                              </div>
                            )}
                          </td>

                          {/* Direct Report Action Column */}
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
                              className="inline-flex items-center gap-1 bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 px-2 py-1 rounded text-[11px] font-semibold"
                            >
                              <History size={11} /> History
                            </button>
                          </td>

                          <td className="border border-slate-300 px-2.5 py-1.5 text-center">
                            <button
                              type="button"
                              onClick={() => handleEditClick(visit)}
                              className="inline-flex items-center gap-1 bg-amber-50 text-amber-700 border border-amber-200 hover:bg-amber-100 px-2 py-1 rounded text-[11px] font-semibold"
                            >
                              <Edit3 size={11} /> Edit
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
          MODAL 1: BEFORE SALE REPORT (WITH FEED CHANGE AGENT BOX)
      ================================================== */}
      {showBeforeSaleModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white p-5 rounded-xl w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center mb-3 border-b pb-2.5">
              <h3 className="text-base font-bold text-slate-800">
                Before Sale Report —{" "}
                {activeReportVisit.batch_label ||
                  activeReportVisit.farmers?.farmer_name}
              </h3>
              <button onClick={() => setShowBeforeSaleModal(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="mb-3 bg-slate-50 p-2.5 rounded-lg border flex items-center justify-between text-xs">
              <div>
                <span className="font-bold text-slate-600 block">
                  Visit Date
                </span>
                <input
                  type="date"
                  value={reportVisitDate}
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="mt-1 border rounded px-2 py-1"
                />
              </div>
              <div className="text-right">
                <span className="text-slate-500">Age on Visit Date:</span>
                <div className="font-bold text-emerald-700 text-sm">
                  {calculateAge(
                    activeReportVisit.hatch_date,
                    activeReportVisit.category,
                    reportVisitDate
                  )}
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 text-xs">
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
                  className="w-full border rounded px-2.5 py-1.5"
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
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed (T. Kg)
                </label>
                <input
                  type="number"
                  value={beforeForm.feed_t_kg}
                  onChange={(e) =>
                    updateBeforeCalculations({ feed_t_kg: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed (g/b) [Auto]
                </label>
                <input
                  type="number"
                  disabled
                  value={beforeForm.feed_gm_b}
                  className="w-full border rounded px-2.5 py-1.5 bg-slate-100 font-semibold"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  FCR (W.out M) [Auto]
                </label>
                <input
                  type="number"
                  disabled
                  value={beforeForm.fcr_w_out_m}
                  className="w-full border rounded px-2.5 py-1.5 bg-slate-100 font-semibold"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  FCR (With M) [Auto]
                </label>
                <input
                  type="number"
                  disabled
                  value={beforeForm.fcr_with_m}
                  className="w-full border rounded px-2.5 py-1.5 bg-slate-100 font-semibold"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Uneven
                </label>
                <input
                  type="text"
                  value={beforeForm.uneven}
                  onChange={(e) =>
                    setBeforeForm({ ...beforeForm, uneven: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Indigestion
                </label>
                <input
                  type="text"
                  value={beforeForm.indigestion}
                  onChange={(e) =>
                    setBeforeForm({
                      ...beforeForm,
                      indigestion: e.target.value,
                    })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Orange Feces
                </label>
                <input
                  type="text"
                  value={beforeForm.orange_feces}
                  onChange={(e) =>
                    setBeforeForm({
                      ...beforeForm,
                      orange_feces: e.target.value,
                    })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Watery Feces
                </label>
                <input
                  type="text"
                  value={beforeForm.watery_feces}
                  onChange={(e) =>
                    setBeforeForm({
                      ...beforeForm,
                      watery_feces: e.target.value,
                    })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
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
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Ventilation
                </label>
                <input
                  type="text"
                  value={beforeForm.ventilation}
                  onChange={(e) =>
                    setBeforeForm({
                      ...beforeForm,
                      ventilation: e.target.value,
                    })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>

              {/* Feed Company Dropdown */}
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
                  className="w-full border border-emerald-400 rounded px-2.5 py-1.5 bg-white font-semibold"
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
                <input
                  type="text"
                  value={beforeForm.feed_type}
                  onChange={(e) =>
                    setBeforeForm({ ...beforeForm, feed_type: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
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
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>

              {/* IF FEED COMPANY CHANGED -> SHOW AGENT ASSIGNMENT BOX */}
              {beforeChangedComp && (
                <div className="col-span-2 sm:col-span-3 bg-amber-50 border border-amber-300 rounded-xl p-3 space-y-2">
                  <div className="text-xs font-bold text-amber-900">
                    🔄 Feed Company Changed to{" "}
                    <span className="underline">
                      {beforeChangedComp.company_name}
                    </span>
                    ! Select Agent under {beforeChangedComp.company_name}:
                  </div>
                  <select
                    value={transferAgentId}
                    onChange={(e) => setTransferAgentId(e.target.value)}
                    className="w-full border border-amber-400 rounded-lg px-3 py-1.5 bg-white text-xs font-semibold"
                  >
                    <option value="">
                      -- Select Agent under {beforeChangedComp.company_name} --
                    </option>
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
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-4 border-t mt-4">
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
                Save Before Sale Report
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 2: AFTER SALE REPORT (WITH "SOLD" BUTTON)
      ================================================== */}
      {showAfterSaleModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white p-5 rounded-xl w-full max-w-md max-h-[90vh] overflow-y-auto shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center mb-3 border-b pb-2.5">
              <h3 className="text-base font-bold text-slate-800">
                After Sale Report —{" "}
                {activeReportVisit.batch_label ||
                  activeReportVisit.farmers?.farmer_name}
              </h3>
              <button onClick={() => setShowAfterSaleModal(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="mb-3 bg-slate-50 p-2.5 rounded-lg border flex items-center justify-between text-xs">
              <span className="font-bold text-slate-600">Sale / Visit Date</span>
              <input
                type="date"
                value={reportVisitDate}
                onChange={(e) => setReportVisitDate(e.target.value)}
                className="border rounded px-2 py-1"
              />
            </div>

            <div className="space-y-2.5 text-xs">
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Mortality (Pcs)
                </label>
                <input
                  type="number"
                  value={afterForm.mortality}
                  onChange={(e) =>
                    updateAfterCalculations({ mortality: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Average Weight (g)
                </label>
                <input
                  type="number"
                  value={afterForm.weight}
                  onChange={(e) =>
                    updateAfterCalculations({ weight: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
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
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block font-bold text-slate-600 mb-1">
                    FCR (W.out M) [Auto]
                  </label>
                  <input
                    type="number"
                    disabled
                    value={afterForm.fcr_w_out_m}
                    className="w-full border rounded px-2.5 py-1.5 bg-slate-100 font-semibold"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-600 mb-1">
                    FCR (With M) [Auto]
                  </label>
                  <input
                    type="number"
                    disabled
                    value={afterForm.fcr_with_m}
                    className="w-full border rounded px-2.5 py-1.5 bg-slate-100 font-semibold"
                  />
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed Company (Feed Mill)
                </label>
                <select
                  value={afterForm.feed_mill}
                  onChange={(e) => {
                    setAfterForm({ ...afterForm, feed_mill: e.target.value });
                    setTransferAgentId("");
                  }}
                  className="w-full border rounded px-2.5 py-1.5 bg-white"
                >
                  <option value="">Select Feed Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.company_name}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
              </div>

              {afterChangedComp && (
                <div className="bg-amber-50 border border-amber-300 rounded-xl p-2.5 space-y-1.5">
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

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed Name / Type
                </label>
                <input
                  type="text"
                  value={afterForm.feed}
                  onChange={(e) =>
                    setAfterForm({ ...afterForm, feed: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Remarks
                </label>
                <input
                  type="text"
                  value={afterForm.remarks}
                  onChange={(e) =>
                    setAfterForm({ ...afterForm, remarks: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-4 border-t mt-4">
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
                ✓ Sold (Save & Close Batch)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 3: LAYER PERFORMANCE REPORT (WITH "CULLED" BUTTON)
      ================================================== */}
      {showLayerModal && activeReportVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white p-5 rounded-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto shadow-2xl border border-slate-300">
            <div className="flex justify-between items-center mb-3 border-b pb-2.5">
              <h3 className="text-base font-bold text-indigo-900">
                Layer Performance Report —{" "}
                {activeReportVisit.batch_label ||
                  activeReportVisit.farmers?.farmer_name}
              </h3>
              <button onClick={() => setShowLayerModal(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="grid grid-cols-2 gap-3 text-xs">
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Visit Date
                </label>
                <input
                  type="date"
                  value={reportVisitDate}
                  onChange={(e) => setReportVisitDate(e.target.value)}
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  B. Weight (g)
                </label>
                <input
                  type="number"
                  value={layerForm.b_weight_g}
                  onChange={(e) =>
                    setLayerForm({ ...layerForm, b_weight_g: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Feed Intake (g)
                </label>
                <input
                  type="number"
                  value={layerForm.feed_intake_g}
                  onChange={(e) =>
                    setLayerForm({
                      ...layerForm,
                      feed_intake_g: e.target.value,
                    })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
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
                  className="w-full border rounded px-2.5 py-1.5"
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
                  className="w-full border rounded px-2.5 py-1.5"
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
                  className="w-full border rounded px-2.5 py-1.5 bg-white font-semibold"
                >
                  <option value="">Select Feed Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.company_name}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
              </div>

              {layerChangedComp && (
                <div className="col-span-2 bg-amber-50 border border-amber-300 rounded-xl p-2.5 space-y-1.5">
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

              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Disease
                </label>
                <input
                  type="text"
                  value={layerForm.disease}
                  onChange={(e) =>
                    setLayerForm({ ...layerForm, disease: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-600 mb-1">
                  Vaccine
                </label>
                <input
                  type="text"
                  value={layerForm.vaccine}
                  onChange={(e) =>
                    setLayerForm({ ...layerForm, vaccine: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
              <div className="col-span-2">
                <label className="block font-bold text-slate-600 mb-1">
                  Remarks
                </label>
                <input
                  type="text"
                  value={layerForm.remarks}
                  onChange={(e) =>
                    setLayerForm({ ...layerForm, remarks: e.target.value })
                  }
                  className="w-full border rounded px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="flex flex-wrap justify-between items-center gap-2 pt-4 border-t mt-4">
              <button
                type="button"
                onClick={() => handleLayerReportSubmit(true)}
                className="px-4 py-1.5 rounded-lg bg-red-600 hover:bg-red-700 text-white text-xs font-bold shadow-sm"
              >
                Culled (Save & Close Layer Batch)
              </button>

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
                  Save Visit Report
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 4: VACCINE SCHEDULE MANAGER (BROILER, LAYER, SONALI)
      ================================================== */}
      {showVaccineModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white rounded-xl p-5 w-full max-w-3xl max-h-[90vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-4">
            <div className="flex justify-between items-center border-b pb-2.5">
              <h3 className="text-base font-bold text-slate-800">
                💉 Vaccine Schedule (Broiler, Layer, Sonali)
              </h3>
              <button onClick={() => setShowVaccineModal(false)}>
                <X size={18} />
              </button>
            </div>

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
                  className="bg-emerald-600 text-white px-2.5 rounded font-bold shrink-0"
                >
                  <Plus size={14} />
                </button>
              </div>
            </form>

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
                    <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                      Delete
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {vaccines.map((v) => (
                    <tr key={v.id} className="hover:bg-slate-50">
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
                      <td className="border border-slate-300 px-2.5 py-1 text-center">
                        <button
                          type="button"
                          onClick={() => handleDeleteVaccine(v.id)}
                          className="text-red-500 hover:text-red-700"
                        >
                          <Trash2 size={13} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          MODAL 5: CLEAN TABLE HISTORY MODAL (CATEGORY SPECIFIC)
      ================================================== */}
      {showHistoryModal && selectedHistoryVisit && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white p-5 rounded-xl w-full max-w-5xl max-h-[90vh] overflow-y-auto shadow-2xl border border-slate-300 space-y-5">
            <div className="flex justify-between items-center border-b pb-2.5">
              <div>
                <h3 className="text-base font-bold text-slate-800">
                  History —{" "}
                  {selectedHistoryVisit.batch_label ||
                    selectedHistoryVisit.farmers?.farmer_name}{" "}
                  ({selectedHistoryVisit.category})
                </h3>
                <p className="text-xs text-slate-500">
                  Hatch Date: {selectedHistoryVisit.hatch_date || "-"} • Breed:{" "}
                  {selectedHistoryVisit.breed || "-"} • Birds:{" "}
                  {selectedHistoryVisit.t_birds || 0}
                </p>
              </div>
              <button onClick={() => setShowHistoryModal(false)}>
                <X size={18} />
              </button>
            </div>

            {(selectedHistoryVisit.category || "").toLowerCase() === "layer" ? (
              /* ONLY LAYER HISTORY TABLE FOR LAYER BATCHES */
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
                          Age
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
                      {farmerHistory.layer.length === 0 ? (
                        <tr>
                          <td
                            colSpan={9}
                            className="border border-slate-300 py-4 text-center text-slate-400"
                          >
                            No layer reports found for this batch.
                          </td>
                        </tr>
                      ) : (
                        farmerHistory.layer.map((rep: any, i: number) => (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="border border-slate-300 px-2 py-1">
                              {rep.visit_date || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 font-medium">
                              {calculateAge(
                                selectedHistoryVisit.hatch_date,
                                "Layer",
                                rep.visit_date
                              )}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {rep.b_weight_g}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {rep.feed_intake_g}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold text-indigo-700">
                              {rep.egg_production}%
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {rep.egg_weight}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {rep.feed_company || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {[rep.disease, rep.vaccine]
                                .filter(Boolean)
                                .join(" / ") || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {rep.remarks || "-"}
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              /* ONLY BEFORE & AFTER SALE TABLES FOR BROILER / SONALI BATCHES */
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
                            Age
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
                            FCR (w/o M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            FCR (With M)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed Company
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Remarks
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {farmerHistory.before.length === 0 ? (
                          <tr>
                            <td
                              colSpan={9}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No before sale reports found for this batch.
                            </td>
                          </tr>
                        ) : (
                          farmerHistory.before.map((rep: any, i: number) => (
                            <tr key={i} className="hover:bg-slate-50">
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.visit_date}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {calculateAge(
                                  selectedHistoryVisit.hatch_date,
                                  selectedHistoryVisit.category,
                                  rep.visit_date
                                )}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {rep.mortality_pcs}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {rep.weight_gm_b}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {rep.feed_t_kg}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                {rep.fcr_w_out_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                {rep.fcr_with_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.feed_company || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.remarks || "-"}
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
                            Mortality
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Weight (g)
                          </th>
                          <th className="border border-slate-300 px-2 py-1.5">
                            Feed (Kg)
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
                            Remarks
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {farmerHistory.after.length === 0 ? (
                          <tr>
                            <td
                              colSpan={8}
                              className="border border-slate-300 py-4 text-center text-slate-400"
                            >
                              No after sale reports found for this batch.
                            </td>
                          </tr>
                        ) : (
                          farmerHistory.after.map((rep: any, i: number) => (
                            <tr key={i} className="hover:bg-slate-50">
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.visit_date}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {rep.mortality}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {rep.weight}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center">
                                {rep.feed_consumption_kg}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                {rep.fcr_w_out_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                                {rep.fcr_with_m}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.feed_mill || "-"}
                              </td>
                              <td className="border border-slate-300 px-2 py-1">
                                {rep.remarks || "-"}
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

      {/* ==================================================
          MODAL 6: EDIT VISIT MODAL
      ================================================== */}
      {showEditModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-3">
          <div className="bg-white p-5 rounded-xl w-full max-w-md shadow-2xl border border-slate-300 space-y-3 text-xs">
            <div className="flex justify-between items-center border-b pb-2">
              <h3 className="text-sm font-bold text-slate-800">
                Edit Batch Details
              </h3>
              <button onClick={() => setShowEditModal(false)}>
                <X size={18} />
              </button>
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