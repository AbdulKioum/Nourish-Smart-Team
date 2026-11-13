import { useEffect, useState, useRef } from "react";
import LocationPicker from "../../components/maps/LocationPicker";
import { supabase } from "../../lib/supabase";
import { getAgents } from "../../services/agentService";
import {
  getRetailers,
  createRetailer,
  updateRetailer,
  deleteRetailer,
} from "../../services/retailerService";
import { useAuth } from "../../context/AuthContext";
import {
  MapContainer,
  Marker,
  Popup,
  Tooltip,
  useMap,
} from "react-leaflet";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import { saveAs } from "file-saver";
import { getBazars } from "../../services/bazarService";
import { createRetailerMarker } from "../../utils/mapMarkers";
import MapLayers from "../../components/maps/MapLayers";
import ExcelJS from "exceljs";
import {
  MapPin,
  Plus,
  Search,
  RotateCcw,
  FileSpreadsheet,
  BarChart3,
  Map as MapIcon,
  X,
  CheckCircle2,
  Navigation,
} from "lucide-react";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Capacitor } from "@capacitor/core";

delete (L.Icon.Default.prototype as any)._getIconUrl;

L.Icon.Default.mergeOptions({
  iconRetinaUrl:
    "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png",
  iconUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png",
  shadowUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png",
});

const EMPTY_FORM = {
  agent_id: "",
  company_id: "",
  retailer_name: "",
  proprietor: "",
  mobile: "",
  bazar: "",
  union_name: "",
  upazila: "",
  district: "",
  broiler_sale: "0",
  layer_sale: "0",
  sonali_sale: "0",
  cattle_sale: "0",
  fish_sale: "0",
  monthly_sale: "0",
  focus_item: "",
  latitude: "",
  longitude: "",
  status: "Active",
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

function getSaleBreakdown(r: any) {
  const items = [
    { label: "Broiler", val: Number(r?.broiler_sale || 0) },
    { label: "Layer", val: Number(r?.layer_sale || 0) },
    { label: "Sonali", val: Number(r?.sonali_sale || 0) },
    { label: "Cattle", val: Number(r?.cattle_sale || 0) },
    { label: "Fish", val: Number(r?.fish_sale || 0) },
  ];
  return items.filter((i) => i.val > 0);
}

function RecenterMap({ position }: { position: [number, number] }) {
  const map = useMap();
  useEffect(() => {
    if (!position) return;
    map.flyTo(position, 17, { animate: true, duration: 1.3 });
  }, [position, map]);
  return null;
}

function FitBounds({ positions }: { positions: [number, number][] }) {
  const map = useMap();
  useEffect(() => {
    if (positions.length > 0) {
      map.fitBounds(positions, { padding: [45, 45] });
    }
  }, [positions, map]);
  return null;
}

export default function RetailerPage() {
  const { role } = useAuth();
  const mapSectionRef = useRef<HTMLDivElement | null>(null);
  const retailerMarkerRefs = useRef<Record<number, L.Marker | null>>({});

  if (role === "viewer") {
    return <div className="p-6 text-center text-slate-500">Access Denied</div>;
  }

  const isFullAccess = role === "super_admin" || role === "manager";

  // Assigned Areas
  
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);

  // Master Data
  const [retailers, setRetailers] = useState<any[]>([]);
  const [agents, setAgents] = useState<any[]>([]);
  const [companies, setCompanies] = useState<any[]>([]);
  const [districts, setDistricts] = useState<any[]>([]);
  const [upazilas, setUpazilas] = useState<any[]>([]);
  const [unions, setUnions] = useState<any[]>([]);
  const [bazars, setBazars] = useState<any[]>([]);

  // Form & Modals
  const [showForm, setShowForm] = useState(false);
  const [showMap, setShowMap] = useState(false);
  const [showRetailerMap, setShowRetailerMap] = useState(false);
  const [showSummary, setShowSummary] = useState(false);
  const [selectedRetailer, setSelectedRetailer] = useState<any>(null);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [form, setForm] = useState({ ...EMPTY_FORM });

  // Location Picker Live Location State
  const [pickerCenterOverride, setPickerCenterOverride] = useState<
    [number, number] | null
  >(null);
  const [locatingInPicker, setLocatingInPicker] = useState(false);

  // Filters
  const [search, setSearch] = useState("");
  const [companyFilter, setCompanyFilter] = useState("");
  const [agentFilter, setAgentFilter] = useState("");
  const [districtFilter, setDistrictFilter] = useState("");
  const [upazilaFilter, setUpazilaFilter] = useState("");
  const [unionFilter, setUnionFilter] = useState("");
  const [bazarFilter, setBazarFilter] = useState("");

  const [filterUpazilas, setFilterUpazilas] = useState<any[]>([]);
  const [filterUnions, setFilterUnions] = useState<any[]>([]);
  const [filterBazars, setFilterBazars] = useState<any[]>([]);

  // Main Map States
  const [userLocation, setUserLocation] = useState<[number, number] | null>(
    null
  );
  const [nearestRetailers, setNearestRetailers] = useState<any[]>([]);
  const [selectedLocation, setSelectedLocation] = useState<
    [number, number] | null
  >(null);
  const [mapCenter] = useState<[number, number]>([24.75, 90.4]);
  const [myLocationMode, setMyLocationMode] = useState(false);

  // ==================================================
  // INITIALIZE USER & DATA (STRICT AREA RESTRICTION)
  // ==================================================
  async function loadData() {
    const {
      data: { user },
    } = await supabase.auth.getUser();

    let uDistricts: string[] = [];
    let uUpazilas: string[] = [];

    if (user) {
      const { data: profile } = await supabase
        .from("users")
        .select("role, districts, upazilas")
        .eq("id", user.id)
        .single();

      if (profile) {
        uDistricts = parseList(profile.districts);
        uUpazilas = parseList(profile.upazilas);
      }
    }


    setAssignedUpazilas(uUpazilas);

    const [
      retailerData,
      agentData,
      bazarData,
      { data: companyData },
      { data: districtData },
    ] = await Promise.all([
      getRetailers(),
      getAgents(),
      getBazars(),
      supabase.from("companies").select("*").order("company_name"),
      supabase.from("districts").select("*").order("name"),
    ]);

    const allRetailers = retailerData || [];
    const allowedRetailers = isFullAccess
      ? allRetailers
      : allRetailers.filter((r: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(r.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(r.district);
          return matchUpa && matchDist;
        });

    setRetailers(allowedRetailers);

    // Filter Parent Agents strictly by user's assigned districts/upazilas
    const allAgents = agentData || [];
    const allowedAgents = isFullAccess
      ? allAgents
      : allAgents.filter((a: any) => {
          const aDist = a.district || a.district_name || "";
          const aUpa = a.upazila || a.upazila_name || "";
          const matchUpa =
            uUpazilas.length === 0 || !aUpa || uUpazilas.includes(aUpa);
          const matchDist =
            uDistricts.length === 0 || !aDist || uDistricts.includes(aDist);
          const hasAllowedRetailer = allowedRetailers.some(
            (r: any) => String(r.agent_id) === String(a.id)
          );
          if (aUpa || aDist) return matchUpa && matchDist;
          return hasAllowedRetailer;
        });

    setAgents(allowedAgents);
    setBazars(bazarData || []);
    setCompanies(companyData || []);

    const allDistricts = districtData || [];
    const allowedDists = isFullAccess
      ? allDistricts
      : allDistricts.filter(
          (d: any) => uDistricts.length === 0 || uDistricts.includes(d.name)
        );

    setDistricts(allowedDists);

    if (!districtFilter) {
      if (allowedDists.length === 1) {
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

  // Auto calculate Total Monthly Sale in Form
  useEffect(() => {
    const total =
      Number(form.broiler_sale || 0) +
      Number(form.layer_sale || 0) +
      Number(form.sonali_sale || 0) +
      Number(form.cattle_sale || 0) +
      Number(form.fish_sale || 0);

    setForm((prev) => ({
      ...prev,
      monthly_sale: total.toString(),
    }));
  }, [
    form.broiler_sale,
    form.layer_sale,
    form.sonali_sale,
    form.cattle_sale,
    form.fish_sale,
  ]);

  function clearFilters() {
    setCompanyFilter("");
    setAgentFilter("");
    setDistrictFilter(districts.length === 1 ? districts[0].name : "");
    setUpazilaFilter(
      !isFullAccess && assignedUpazilas.length === 1 ? assignedUpazilas[0] : ""
    );
    setUnionFilter("");
    setBazarFilter("");
    setSearch("");
  }

  // Jump to Retailer on Map when clicking "View on Map" button in table row
  function handleViewRetailerOnMap(r: any) {
    if (!r.latitude || !r.longitude) {
      alert("Location not available for this retailer.");
      return;
    }
    const lat = Number(r.latitude);
    const lng = Number(r.longitude);
    setSelectedRetailer(null);
    setShowSummary(false);
    setShowMap(true);
    setMyLocationMode(false);
    setSelectedLocation([lat, lng]);

    setTimeout(() => {
      mapSectionRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "center",
      });
      const marker = retailerMarkerRefs.current[r.id];
      if (marker) {
        marker.openPopup();
      }
    }, 350);
  }

  // ==================================================
  // DEPENDENT LOCATION DROPDOWNS (RESPECTING ASSIGNMENTS)
  // ==================================================
  const selectedDistrict = districts.find((d) => d.name === form.district);

  useEffect(() => {
    if (!form.district) {
      setUpazilas([]);
      return;
    }
    async function loadUpazilas() {
      const district = districts.find((d) => d.name === form.district);
      if (!district) return;

      const { data } = await supabase
        .from("upazilas")
        .select("*")
        .eq("district_id", district.id)
        .order("name");

      const list = data || [];
      const allowed = isFullAccess
        ? list
        : list.filter(
            (u: any) =>
              assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
          );
      setUpazilas(allowed);
    }
    loadUpazilas();
  }, [form.district, districts, isFullAccess, assignedUpazilas]);

  const selectedUpazila = upazilas.find((u: any) => u.name === form.upazila);

  useEffect(() => {
    if (!form.upazila) {
      setUnions([]);
      return;
    }
    async function loadUnions() {
      const upazila = upazilas.find((u) => u.name === form.upazila);
      if (!upazila) return;

      const { data } = await supabase
        .from("unions")
        .select("*")
        .eq("upazila_id", upazila.id)
        .order("name");

      setUnions(data || []);
    }
    loadUnions();
  }, [form.upazila, upazilas]);

  const selectedUnion = unions.find((u: any) => u.name === form.union_name);
  const filteredBazars = bazars.filter(
    (b: any) => b.union_id === selectedUnion?.id
  );

  // Filter Bar Dependent Dropdowns
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
  }, [upazilaFilter, filterUpazilas]);

  useEffect(() => {
    if (!unionFilter) {
      setFilterBazars([]);
      return;
    }
    async function loadFilterBazars() {
      const union = filterUnions.find((u) => u.name === unionFilter);
      if (!union) return;

      const { data } = await supabase
        .from("bazars")
        .select("*")
        .eq("union_id", union.id)
        .order("name");

      setFilterBazars(data || []);
    }
    loadFilterBazars();
  }, [unionFilter, filterUnions]);

  // Strictly Filter Parent Agents in Filter Bar by Company + Selected District/Upazila + User Area
  const availableFilterAgents = agents.filter((a: any) => {
    if (companyFilter && String(a.company_id) !== String(companyFilter)) {
      return false;
    }
    const aDist = a.district || a.district_name || "";
    const aUpa = a.upazila || a.upazila_name || "";

    if (districtFilter) {
      if (aDist && aDist !== districtFilter) return false;
      if (!aDist) {
        const hasRetailerInDist = retailers.some(
          (r) =>
            String(r.agent_id) === String(a.id) &&
            r.district === districtFilter
        );
        if (!hasRetailerInDist) return false;
      }
    }

    if (upazilaFilter) {
      if (aUpa && aUpa !== upazilaFilter) return false;
      if (!aUpa) {
        const hasRetailerInUpa = retailers.some(
          (r) =>
            String(r.agent_id) === String(a.id) && r.upazila === upazilaFilter
        );
        if (!hasRetailerInUpa) return false;
      }
    }

    return true;
  });

  // Parent Agents inside Add/Edit Form (Filtered by selected Company)
 // const formAvailableAgents = agents.filter((a: any) => {
  //  if (form.company_id && String(a.company_id) !== String(form.company_id)) {
   //   return false;
   // }
 //   return true;
 // });

  // ==================================================
  // SAVE, DELETE & LOCATE ME
  // ==================================================
  async function handleSave() {
    try {
      if (
        !form.company_id ||
        !form.retailer_name.trim() ||
        !form.district ||
        !form.upazila
      ) {
        alert(
          "Company, Retailer Name, District and Upazila are required."
        );
        return;
      }

      const payload = {
        ...form,
        agent_id: form.agent_id ? Number(form.agent_id) : null,
        company_id: Number(form.company_id),
        broiler_sale: Number(form.broiler_sale || 0),
        layer_sale: Number(form.layer_sale || 0),
        sonali_sale: Number(form.sonali_sale || 0),
        cattle_sale: Number(form.cattle_sale || 0),
        fish_sale: Number(form.fish_sale || 0),
        monthly_sale: form.monthly_sale === "" ? 0 : Number(form.monthly_sale),
        latitude: form.latitude === "" ? null : Number(form.latitude),
        longitude: form.longitude === "" ? null : Number(form.longitude),
      };

      if (editingId) {
        await updateRetailer(editingId, payload);
        alert("Retailer Updated");
      } else {
        await createRetailer(payload);
        alert("Retailer Saved");
      }

      setShowForm(false);
      setEditingId(null);
      setForm({ ...EMPTY_FORM });
      loadData();
    } catch (error: any) {
      alert(error?.message || "Save Failed");
    }
  }

  async function handleDelete(id: number) {
    if (!confirm("Delete Retailer?")) return;
    await deleteRetailer(id);
    loadData();
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
        setForm((prev) => ({
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

  function getDistanceKm(
    lat1: number,
    lon1: number,
    lat2: number,
    lon2: number
  ) {
    const R = 6371;
    const dLat = ((lat2 - lat1) * Math.PI) / 180;
    const dLon = ((lon2 - lon1) * Math.PI) / 180;
    const a =
      Math.sin(dLat / 2) * Math.sin(dLat / 2) +
      Math.cos((lat1 * Math.PI) / 180) *
        Math.cos((lat2 * Math.PI) / 180) *
        Math.sin(dLon / 2) *
        Math.sin(dLon / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return Math.round(R * c * 100) / 100;
  }

  function getMyLocation() {
    if (!navigator.geolocation) {
      alert("Geolocation not supported");
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (position) => {
        const coords: [number, number] = [
          position.coords.latitude,
          position.coords.longitude,
        ];
        setUserLocation(coords);
        setSelectedLocation(coords);
        setMyLocationMode(true);

        const nearest = filteredRetailers
          .filter((a) => a.latitude && a.longitude)
          .map((a) => ({
            ...a,
            distance: getDistanceKm(
              coords[0],
              coords[1],
              Number(a.latitude),
              Number(a.longitude)
            ),
          }))
          .sort((a, b) => a.distance - b.distance)
          .slice(0, 10);

        setNearestRetailers(nearest);
      },
      () => {
        alert("Location access denied");
      }
    );
  }

  // ==================================================
  // FILTERED RETAILERS & COMPANY-WISE GROUPING
  // ==================================================
  const filteredRetailers = retailers.filter((r) => {
    const q = search.toLowerCase();
    const matchSearch =
      !q ||
      r.retailer_name?.toLowerCase().includes(q) ||
      r.proprietor?.toLowerCase().includes(q) ||
      r.mobile?.toLowerCase().includes(q) ||
      r.bazar?.toLowerCase().includes(q);

    const matchCompany =
      companyFilter === "" || String(r.company_id) === companyFilter;
    const matchAgent =
      agentFilter === "" || String(r.agent_id) === agentFilter;
    const matchDistrict =
      districtFilter === "" || r.district === districtFilter;
    const matchUpazila =
      upazilaFilter === "" || r.upazila === upazilaFilter;
    const matchUnion = unionFilter === "" || r.union_name === unionFilter;
    const matchBazar = bazarFilter === "" || r.bazar === bazarFilter;

    return (
      matchSearch &&
      matchCompany &&
      matchAgent &&
      matchDistrict &&
      matchUpazila &&
      matchUnion &&
      matchBazar
    );
  });

  const sortedRetailers = [...filteredRetailers].sort((a: any, b: any) => {
    const companyCompare = (
      a.companies?.company_name || ""
    ).localeCompare(b.companies?.company_name || "");
    if (companyCompare !== 0) return companyCompare;
    return (a.retailer_name || "").localeCompare(b.retailer_name || "");
  });

  // Group Retailers by Company
  const companyGroupedRetailers = Object.values(
    sortedRetailers.reduce((acc: Record<string, any>, r: any) => {
      const key = String(r.company_id || "unassigned");
      if (!acc[key]) {
        acc[key] = {
          companyId: key,
          companyName: r.companies?.company_name || "Unassigned Company",
          logoUrl: r.companies?.logo_url || "",
          totalSale: 0,
          retailers: [],
        };
      }
      acc[key].retailers.push(r);
      acc[key].totalSale += Number(r.monthly_sale || 0);
      return acc;
    }, {})
  );

  const selectedCompanyName =
    companies.find((c) => String(c.id) === companyFilter)?.company_name ||
    "All Companies";

  const totalMonthlySale = filteredRetailers.reduce(
    (sum, r) => sum + Number(r.monthly_sale || 0),
    0
  );

  const companySummary = Object.values(
    filteredRetailers.reduce((acc: any, r) => {
      const company = r.companies?.company_name || "Unknown";
      if (!acc[company]) {
        acc[company] = {
          name: company,
          agents: 0,
          broiler: 0,
          layer: 0,
          sonali: 0,
          cattle: 0,
          fish: 0,
          sale: 0,
        };
      }
      acc[company].agents += 1;
      acc[company].broiler += Number(r.broiler_sale || 0);
      acc[company].layer += Number(r.layer_sale || 0);
      acc[company].sonali += Number(r.sonali_sale || 0);
      acc[company].cattle += Number(r.cattle_sale || 0);
      acc[company].fish += Number(r.fish_sale || 0);
      acc[company].sale += Number(r.monthly_sale || 0);
      return acc;
    }, {})
  );

  // ==================================================
  // EXCEL EXPORT
  // ==================================================
 
  async function exportRetailersExcel() {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Retailers");

    worksheet.getRow(1).values = [
      "Company",
      "Parent Agent",
      "Retailer Name",
      "Proprietor",
      "Mobile",
      "Bazar",
      "District",
      "Upazila",
      "Union",
      "Broiler Sale",
      "Layer Sale",
      "Sonali Sale",
      "Cattle Sale",
      "Fish Sale",
      "Total Monthly Sale",
      "Focus Item",
      "Status",
    ];

    sortedRetailers.forEach((r: any) => {
      worksheet.addRow([
        r.companies?.company_name || "",
        r.agents?.agent_name || "",
        r.retailer_name || "",
        r.proprietor || "",
        r.mobile || "",
        r.bazar || "",
        r.district || "",
        r.upazila || "",
        r.union_name || "",
        Number(r.broiler_sale || 0),
        Number(r.layer_sale || 0),
        Number(r.sonali_sale || 0),
        Number(r.cattle_sale || 0),
        Number(r.fish_sale || 0),
        Number(r.monthly_sale || 0),
        r.focus_item || "",
        r.status || "",
      ]);
    });

    for (let col = 1; col <= 17; col++) {
      const cell = worksheet.getCell(1, col);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb:
            col === 15
              ? "15803D"
              : col >= 10 && col <= 14
              ? "D97706"
              : "1E3A8A",
        },
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
    }

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber >= 2) {
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

    for (let row = 2; row <= worksheet.rowCount; row++) {
      for (let col = 10; col <= 15; col++) {
        worksheet.getCell(row, col).numFmt = "#,##0";
        worksheet.getCell(row, col).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
    }

    const widths = [
      24, 22, 22, 18, 16, 18, 16, 16, 16, 14, 14, 14, 14, 14, 18, 18, 14,
    ];
    widths.forEach((w, idx) => (worksheet.getColumn(idx + 1).width = w));
    worksheet.getRow(1).height = 30;
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.autoFilter = { from: "A1", to: "Q1" };
    worksheet.pageSetup.orientation = "landscape";
    worksheet.pageSetup.fitToPage = true;
    worksheet.pageSetup.fitToWidth = 1;
    worksheet.pageSetup.fitToHeight = 0;

    const today = new Date().toISOString().split("T")[0];
    const safeCompanyName = (selectedCompanyName || "Report").replace(/[\s\/\\:@?*]+/g, "_");
    const fileName = `Retailers_${safeCompanyName}_${today}.xlsx`;

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
      saveAs(blob, fileName);
    }
  } catch (error) {
    console.error("Retailer Excel Export Error:", error);
    alert("Retailer Excel export failed.");
  }
}

  function openEditForm(r: any) {
    setEditingId(r.id);
    setForm({
      agent_id: String(r.agent_id || ""),
      company_id: String(r.company_id || ""),
      retailer_name: r.retailer_name || "",
      proprietor: r.proprietor || "",
      mobile: r.mobile || "",
      bazar: r.bazar || "",
      union_name: r.union_name || "",
      upazila: r.upazila || "",
      district: r.district || "",
      broiler_sale: String(r.broiler_sale || 0),
      layer_sale: String(r.layer_sale || 0),
      sonali_sale: String(r.sonali_sale || 0),
      cattle_sale: String(r.cattle_sale || 0),
      fish_sale: String(r.fish_sale || 0),
      monthly_sale: String(r.monthly_sale || 0),
      focus_item: r.focus_item || "",
      latitude: String(r.latitude || ""),
      longitude: String(r.longitude || ""),
      status: r.status || "Active",
    });
    setShowForm(true);
  }

  return (
    <div className="p-2.5 sm:p-4 max-w-7xl mx-auto space-y-3">
      {/* ==================================================
          1. COMPACT FILTER BAR AT THE TOP
      ================================================== */}
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2 items-center">
          {/* Search */}
          <div className="col-span-2 sm:col-span-2 lg:col-span-2 relative">
            <Search
              size={13}
              className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Search retailer, proprietor, mobile, bazar..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full rounded-lg border border-slate-300 pl-8 pr-2.5 py-1.5 text-xs outline-none focus:border-blue-500"
            />
          </div>

          {/* District */}
          <select
            value={districtFilter}
            onChange={(e) => {
              setDistrictFilter(e.target.value);
              setUpazilaFilter("");
              setUnionFilter("");
              setBazarFilter("");
              setAgentFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Districts</option>
            {districts.map((d) => (
              <option key={d.id} value={d.name}>
                {d.name}
              </option>
            ))}
          </select>

          {/* Upazila */}
          <select
            value={upazilaFilter}
            onChange={(e) => {
              setUpazilaFilter(e.target.value);
              setUnionFilter("");
              setBazarFilter("");
              setAgentFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Upazilas</option>
            {filterUpazilas.map((u: any) => (
              <option key={u.id} value={u.name}>
                {u.name}
              </option>
            ))}
          </select>

          {/* Union */}
          <select
            value={unionFilter}
            onChange={(e) => {
              setUnionFilter(e.target.value);
              setBazarFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Unions</option>
            {filterUnions.map((u: any) => (
              <option key={u.id} value={u.name}>
                {u.name}
              </option>
            ))}
          </select>

          {/* Bazar */}
          <select
            value={bazarFilter}
            onChange={(e) => setBazarFilter(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Bazars</option>
            {filterBazars.map((b: any) => (
              <option key={b.id} value={b.name}>
                {b.name}
              </option>
            ))}
          </select>

          {/* Company */}
          <select
            value={companyFilter}
            onChange={(e) => {
              setCompanyFilter(e.target.value);
              setAgentFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Companies</option>
            {companies.map((c) => (
              <option key={c.id} value={c.id}>
                {c.company_name}
              </option>
            ))}
          </select>

          {/* Parent Agent (Strictly Filtered by Assigned District/Upazila) + Small Reset Button */}
          <div className="flex items-center gap-1.5">
            <select
              value={agentFilter}
              onChange={(e) => setAgentFilter(e.target.value)}
              className="w-full rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
            >
              <option value="">All Parent Agents</option>
              {availableFilterAgents.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.agent_name}
                </option>
              ))}
            </select>

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

      {/* ==================================================
          2. ADD RETAILER, SUMMARY, EXPORT, VIEW MAP BAR (BELOW FILTER)
      ================================================== */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-3.5 py-2 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => {
              setEditingId(null);
              const defDist =
                districts.length === 1
                  ? districts[0].name
                  : districts.some((d) => d.name === "Mymensingh")
                  ? "Mymensingh"
                  : "";
              const defUpa =
                !isFullAccess && assignedUpazilas.length === 1
                  ? assignedUpazilas[0]
                  : "";
              setForm({
                ...EMPTY_FORM,
                district: defDist,
                upazila: defUpa,
              });
              setShowForm(true);
            }}
            className="inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <Plus size={14} /> Add Retailer
          </button>

          <button
            onClick={() => {
              setShowMap(false);
              setShowSummary(true);
            }}
            className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <BarChart3 size={14} /> Summary
          </button>

          <button
            onClick={exportRetailersExcel}
            className="inline-flex items-center gap-1.5 bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <FileSpreadsheet size={14} /> Export Excel
          </button>

          <button
            onClick={() => {
              if (!showMap) {
                setMyLocationMode(false);
                setUserLocation(null);
                setSelectedLocation(null);
              }
              setShowSummary(false);
              setShowMap(!showMap);
            }}
            className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition ${
              showMap
                ? "bg-red-600 hover:bg-red-700 text-white"
                : "bg-blue-600 hover:bg-blue-700 text-white"
            }`}
          >
            <MapIcon size={14} />
            {showMap ? "Close Map" : "View Map"}
          </button>
        </div>

        <div className="flex items-center gap-3 text-xs font-semibold text-slate-600">
          <span>
            Total Retailers:{" "}
            <strong className="text-slate-900">
              {filteredRetailers.length}
            </strong>
          </span>
          <span>
            Total Sale:{" "}
            <strong className="text-emerald-700">
              {totalMonthlySale.toLocaleString()}
            </strong>
          </span>
        </div>
      </div>

      {/* ==================================================
          3. INTERACTIVE RETAILER MAP VIEW
      ================================================== */}
      {showMap && (
        <div
          ref={mapSectionRef}
          id="retailer-map-section"
          className="bg-white rounded-xl shadow border border-slate-300 p-3"
        >
          <h3 className="text-sm font-bold text-slate-800 mb-2">
            Retailer Locations
          </h3>

          <div className="w-full overflow-hidden rounded-xl border border-slate-300 relative">
            <MapContainer
              center={mapCenter}
              zoom={8}
              style={{ height: "480px", width: "100%" }}
            >
              <MapLayers />

              {!myLocationMode && !selectedLocation && (
                <FitBounds
                  positions={filteredRetailers
                    .filter((a) => a.latitude && a.longitude)
                    .map((a) => [Number(a.latitude), Number(a.longitude)])}
                />
              )}

              {selectedLocation && <RecenterMap position={selectedLocation} />}

              {userLocation && (
                <Marker position={userLocation}>
                  <Tooltip permanent direction="top" offset={[0, -15]}>
                    📍 My Location
                  </Tooltip>
                  <Popup>
                    <div className="max-h-60 overflow-y-auto w-52 text-xs space-y-1.5">
                      <h3 className="font-bold text-slate-800">
                        📍 Nearest Retailers
                      </h3>
                      {nearestRetailers.length === 0 ? (
                        <div>No nearby retailers found.</div>
                      ) : (
                        nearestRetailers.map((r, index) => (
                          <div key={r.id} className="border-b pb-1">
                            <button
                              onClick={() => {
                                setSelectedLocation([
                                  Number(r.latitude),
                                  Number(r.longitude),
                                ]);
                              }}
                              className="text-blue-600 font-bold hover:underline text-left"
                            >
                              {index + 1}. {r.retailer_name}
                            </button>
                            <div className="text-slate-500">
                              {r.distance} km away
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </Popup>
                </Marker>
              )}

              {filteredRetailers
                .filter((r) => r.latitude && r.longitude)
                .map((r) => {
                  const breakdown = getSaleBreakdown(r);

                  return (
                    <Marker
                      key={r.id}
                      ref={(ref) => {
                        retailerMarkerRefs.current[r.id] = ref;
                      }}
                      position={[Number(r.latitude), Number(r.longitude)]}
                      icon={createRetailerMarker(
                        r.companies?.logo_url || "",
                        r.retailer_name || ""
                      )}
                    >
                      <Popup maxWidth={320}>
                        <div className="w-[260px] text-xs space-y-2">
                          <div className="flex items-center gap-2 border-b pb-1.5">
                            {r.companies?.logo_url && (
                              <img
                                src={r.companies.logo_url}
                                alt="logo"
                                className="w-9 h-9 object-contain border rounded p-0.5 bg-white shrink-0"
                              />
                            )}
                            <div>
                              <h3 className="font-bold text-sm text-slate-800 leading-tight">
                                {r.retailer_name}
                              </h3>
                              <p className="text-[11px] text-slate-500">
                                {r.companies?.company_name}
                                {r.agents?.agent_name
                                  ? ` • Agent: ${r.agents.agent_name}`
                                  : ""}
                              </p>
                            </div>
                          </div>

                          <table className="w-full border-collapse border border-slate-200 text-[11px]">
                            <tbody>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Proprietor
                                </td>
                                <td className="border px-2 py-1">
                                  {r.proprietor || "-"}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Mobile
                                </td>
                                <td className="border px-2 py-1">
                                  {r.mobile || "-"}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Address
                                </td>
                                <td className="border px-2 py-1">
                                  {[
                                    r.bazar,
                                    r.union_name,
                                    r.upazila,
                                    r.district,
                                  ]
                                    .filter(Boolean)
                                    .join(", ")}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Sale Breakdown
                                </td>
                                <td className="border px-2 py-1 font-semibold text-blue-700">
                                  {breakdown.length === 0
                                    ? "-"
                                    : breakdown
                                        .map((b) => `${b.label}: ${b.val}`)
                                        .join(", ")}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Monthly Sale
                                </td>
                                <td className="border px-2 py-1 font-bold text-emerald-700">
                                  {Number(r.monthly_sale || 0).toLocaleString()}
                                </td>
                              </tr>
                              {r.focus_item && (
                                <tr>
                                  <td className="border px-2 py-1 font-semibold bg-slate-50">
                                    Focus Item
                                  </td>
                                  <td className="border px-2 py-1">
                                    {r.focus_item}
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

            <button
              onClick={getMyLocation}
              className="absolute top-3 left-1/2 -translate-x-1/2 z-[1000] flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-full shadow-lg px-3.5 py-1.5 text-xs font-semibold transition"
            >
              <MapPin size={15} /> Locate Me
            </button>
          </div>
        </div>
      )}

      {/* ==================================================
          4. SIMPLE BORDERED TABLE GROUPED BY COMPANY
      ================================================== */}
      {companyGroupedRetailers.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-lg p-8 text-center text-xs text-slate-500">
          No retailers found for the selected filter.
        </div>
      ) : (
        <div className="space-y-4">
          {companyGroupedRetailers.map((group: any) => (
            <div
              key={group.companyId}
              className="bg-white border border-slate-300 rounded-lg overflow-hidden shadow-sm"
            >
              {/* Company Group Heading */}
              <div className="bg-slate-100 border-b border-slate-300 px-3 py-2 flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  {group.logoUrl && (
                    <img
                      src={group.logoUrl}
                      alt={group.companyName}
                      className="w-6 h-6 object-contain rounded bg-white border border-slate-200 p-0.5"
                    />
                  )}
                  <span className="text-xs sm:text-sm font-bold text-slate-800">
                    Company: {group.companyName}
                  </span>
                </div>

                <div className="flex items-center gap-2 text-xs">
                  <span className="font-semibold bg-white border border-slate-300 px-2 py-0.5 rounded text-slate-700">
                    Retailers: {group.retailers.length}
                  </span>
                  <span className="font-semibold bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded text-emerald-800">
                    Total Sale: {group.totalSale.toLocaleString()}
                  </span>
                </div>
              </div>

              {/* Simple Bordered Table under Company */}
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-xs">
                  <thead className="bg-slate-50">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        SL
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Parent Agent
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Retailer Name
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Proprietor
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Mobile
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Bazar
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Union
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Upazila
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Sale Details
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-right">
                        Monthly Sale
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Status
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">
                        Actions
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {group.retailers.map((r: any, idx: number) => {
                      const breakdown = getSaleBreakdown(r);
                      return (
                        <tr key={r.id} className="hover:bg-slate-50">
                          <td className="border border-slate-300 px-2 py-1">
                            {idx + 1}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-slate-700">
                            {r.agents?.agent_name || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 font-semibold text-slate-800">
                            {r.retailer_name}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {r.proprietor || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {r.mobile || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {r.bazar || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {r.union_name || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {r.upazila || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {breakdown.length === 0
                              ? "-"
                              : breakdown
                                  .map((b) => `${b.label}: ${b.val}`)
                                  .join(", ")}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-right font-bold text-slate-800">
                            {Number(r.monthly_sale || 0).toLocaleString()}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {r.status}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap space-x-1">
                            <button
                              onClick={() => setSelectedRetailer(r)}
                              className="bg-blue-600 hover:bg-blue-700 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              View
                            </button>
                            <button
                              onClick={() => handleViewRetailerOnMap(r)}
                              className="inline-flex items-center gap-1 bg-teal-600 hover:bg-teal-700 text-white px-2 py-0.5 rounded text-[11px]"
                              >
                                <MapPin size={10} className="shrink-0" />
                                <span>Map</span>
                            </button>
                            <button
                              onClick={() => openEditForm(r)}
                              className="bg-yellow-500 hover:bg-yellow-600 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              Edit
                            </button>
                            <button
                              onClick={() => handleDelete(r.id)}
                              className="bg-red-600 hover:bg-red-700 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              Delete
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
          5. VIEW RETAILER MODAL (CLEAN BORDERED TABLE)
      ================================================== */}
      {selectedRetailer && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white rounded-xl w-full max-w-lg max-h-[90vh] flex flex-col overflow-hidden shadow-2xl border border-slate-300">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-300 bg-slate-100">
              <h2 className="text-sm sm:text-base font-bold text-slate-800">
                Retailer: {selectedRetailer.retailer_name}
              </h2>
              <button
                onClick={() => setSelectedRetailer(null)}
                className="text-slate-500 hover:text-slate-800"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-4 overflow-y-auto">
              <table className="w-full border-collapse border border-slate-300 text-xs">
                <tbody>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold w-1/3">
                      Company
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedRetailer.companies?.company_name || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Parent Agent
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedRetailer.agents?.agent_name || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Retailer Name
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-bold">
                      {selectedRetailer.retailer_name}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Proprietor
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedRetailer.proprietor || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Mobile
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedRetailer.mobile || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Address
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {[
                        selectedRetailer.bazar,
                        selectedRetailer.union_name,
                        selectedRetailer.upazila,
                        selectedRetailer.district,
                      ]
                        .filter(Boolean)
                        .join(", ")}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Sale Breakdown
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-semibold text-blue-700">
                      Broiler: {selectedRetailer.broiler_sale || 0} | Layer:{" "}
                      {selectedRetailer.layer_sale || 0} | Sonali:{" "}
                      {selectedRetailer.sonali_sale || 0} | Cattle:{" "}
                      {selectedRetailer.cattle_sale || 0} | Fish:{" "}
                      {selectedRetailer.fish_sale || 0}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Total Monthly Sale
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-bold text-emerald-700">
                      {Number(
                        selectedRetailer.monthly_sale || 0
                      ).toLocaleString()}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Focus Item
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedRetailer.focus_item || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Status
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedRetailer.status || "Active"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Coordinates
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedRetailer.latitude && selectedRetailer.longitude
                        ? `${selectedRetailer.latitude}, ${selectedRetailer.longitude}`
                        : "-"}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          6. ADD / EDIT RETAILER MODAL
      ================================================== */}
      {showForm && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-5 w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h2 className="text-base font-bold text-slate-800">
                {editingId ? "Edit Retailer" : "Add Retailer"}
              </h2>
              <button
                onClick={() => {
                  setShowForm(false);
                  setEditingId(null);
                }}
              >
                <X size={18} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Company *
                </label>
                <select
                  value={form.company_id}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      company_id: e.target.value,
                      agent_id: "",
                    })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Parent Agent
                </label>
                <select
                  value={form.agent_id}
                  onChange={(e) =>
                    setForm({ ...form, agent_id: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Parent Agent</option>
                  {agents.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.agent_name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Retailer Name *
                </label>
                <input
                  placeholder="Retailer Name"
                  value={form.retailer_name}
                  onChange={(e) =>
                    setForm({ ...form, retailer_name: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Proprietor
                </label>
                <input
                  placeholder="Proprietor"
                  value={form.proprietor}
                  onChange={(e) =>
                    setForm({ ...form, proprietor: e.target.value })
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
                  onChange={(e) =>
                    setForm({ ...form, mobile: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  District *
                </label>
                <select
                  value={form.district}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      district: e.target.value,
                      upazila: "",
                      union_name: "",
                      bazar: "",
                    })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select District</option>
                  {districts.map((d) => (
                    <option key={d.id} value={d.name}>
                      {d.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Upazila *
                </label>
                <select
                  value={form.upazila}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      upazila: e.target.value,
                      union_name: "",
                      bazar: "",
                    })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Upazila</option>
                  {upazilas.map((u) => (
                    <option key={u.id} value={u.name}>
                      {u.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Union
                </label>
                <select
                  value={form.union_name}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      union_name: e.target.value,
                      bazar: "",
                    })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Union</option>
                  {unions.map((u: any) => (
                    <option key={u.id} value={u.name}>
                      {u.name}
                    </option>
                  ))}
                </select>
              </div>

              <div className="sm:col-span-2">
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Bazar
                </label>
                <select
                  value={form.bazar}
                  onChange={(e) =>
                    setForm({ ...form, bazar: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Bazar</option>
                  {filteredBazars.map((bazar) => (
                    <option key={bazar.id} value={bazar.name}>
                      {bazar.name}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Map Location Trigger */}
            <div className="mb-3 flex flex-wrap items-center gap-2 bg-slate-50 border border-slate-200 rounded p-2.5">
              <button
                type="button"
                onClick={() => {
                  setPickerCenterOverride(null);
                  setShowRetailerMap(true);
                }}
                className="inline-flex items-center gap-1.5 bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-semibold"
              >
                <MapPin size={13} /> Select Location from Map
              </button>

              {form.latitude && form.longitude && (
                <span className="text-emerald-700 text-xs font-semibold">
                  📍 Selected ({Number(form.latitude).toFixed(4)},{" "}
                  {Number(form.longitude).toFixed(4)})
                </span>
              )}
            </div>

            {/* Sales Inputs */}
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 mb-3 bg-slate-50 border border-slate-200 rounded p-2.5">
              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Broiler Sale
                </label>
                <input
                  type="number"
                  value={form.broiler_sale}
                  onChange={(e) =>
                    setForm({ ...form, broiler_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Layer Sale
                </label>
                <input
                  type="number"
                  value={form.layer_sale}
                  onChange={(e) =>
                    setForm({ ...form, layer_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Sonali Sale
                </label>
                <input
                  type="number"
                  value={form.sonali_sale}
                  onChange={(e) =>
                    setForm({ ...form, sonali_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Cattle Sale
                </label>
                <input
                  type="number"
                  value={form.cattle_sale}
                  onChange={(e) =>
                    setForm({ ...form, cattle_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Fish Sale
                </label>
                <input
                  type="number"
                  value={form.fish_sale}
                  onChange={(e) =>
                    setForm({ ...form, fish_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-bold text-emerald-800 mb-1">
                  Total Monthly Sale
                </label>
                <input
                  type="number"
                  value={form.monthly_sale}
                  readOnly
                  className="w-full rounded border border-emerald-300 bg-emerald-50 px-2 py-1 text-xs font-bold text-emerald-900"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Focus Item
                </label>
                <select
                  value={form.focus_item}
                  onChange={(e) =>
                    setForm({ ...form, focus_item: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Focus Item</option>
                  <option value="Broiler Feed">Broiler Feed</option>
                  <option value="Layer Feed">Layer Feed</option>
                  <option value="Sonali Feed">Sonali Feed</option>
                  <option value="Cattle Feed">Cattle Feed</option>
                  <option value="Fish Feed">Fish Feed</option>
                  <option value="Goat Feed">Goat Feed</option>
                  <option value="Duck Feed">Duck Feed</option>
                  <option value="Others">Others</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Status
                </label>
                <select
                  value={form.status}
                  onChange={(e) =>
                    setForm({ ...form, status: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="Active">Active</option>
                  <option value="Inactive">Inactive</option>
                </select>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t">
              <button
                onClick={() => {
                  setShowForm(false);
                  setEditingId(null);
                }}
                className="px-4 py-1.5 rounded border border-slate-300 text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleSave}
                className="px-4 py-1.5 rounded bg-green-600 text-white text-xs font-semibold"
              >
                {editingId ? "Update Retailer" : "Save Retailer"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          7. LOCATION PICKER MODAL WITH LOCATE ME BUTTON
      ================================================== */}
      {showRetailerMap && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[60] p-2 sm:p-4">
          <div className="bg-white p-3 sm:p-4 rounded-xl w-full max-w-4xl shadow-2xl space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="font-bold text-sm sm:text-base">
                Select Retailer Location
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

                {form.latitude && form.longitude && (
                  <button
                    type="button"
                    onClick={() => setShowRetailerMap(false)}
                    className="inline-flex items-center gap-1 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                  >
                    <CheckCircle2 size={13} /> Use Selected Location
                  </button>
                )}

                <button onClick={() => setShowRetailerMap(false)}>
                  <X size={18} />
                </button>
              </div>
            </div>

            <LocationPicker
              key={
                pickerCenterOverride
                  ? `${pickerCenterOverride[0]}_${pickerCenterOverride[1]}`
                  : "default_retailer_picker"
              }
              center={
                pickerCenterOverride
                  ? pickerCenterOverride
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
                  : selectedUpazila
                  ? 13
                  : selectedDistrict
                  ? 11
                  : 7
              }
              selectedPosition={
                form.latitude && form.longitude
                  ? [Number(form.latitude), Number(form.longitude)]
                  : null
              }
              onSelect={(lat, lng) => {
                setForm({
                  ...form,
                  latitude: lat.toString(),
                  longitude: lng.toString(),
                });
                setShowRetailerMap(false);
              }}
              markers={
                <>
                  {retailers
                    .filter((a) => a.latitude && a.longitude)
                    .map((r) => (
                      <Marker
                        key={r.id}
                        position={[Number(r.latitude), Number(r.longitude)]}
                        icon={createRetailerMarker(
                          r.companies?.logo_url,
                          r.retailer_name
                        )}
                      />
                    ))}
                </>
              }
            />
          </div>
        </div>
      )}

      {/* ==================================================
          8. RETAILER SUMMARY MODAL
      ================================================== */}
      {showSummary && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-6 w-full max-w-3xl max-h-[85vh] overflow-auto">
            <div className="flex justify-between items-center pb-3 mb-4 border-b">
              <div>
                <h2 className="text-base sm:text-lg font-bold">
                  Retailer Summary — {selectedCompanyName}
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

            <div className="mb-3 flex items-center gap-4 text-xs sm:text-sm font-semibold">
              <span>Total Retailers: {filteredRetailers.length}</span>
              <span className="text-emerald-700">
                Total Monthly Sale: {totalMonthlySale.toLocaleString()}
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-slate-100">
                  <tr>
                    <th className="border px-2.5 py-1.5 text-left">Company</th>
                    <th className="border px-2 py-1.5 text-center">
                      Retailers
                    </th>
                    <th className="border px-2 py-1.5 text-right">Broiler</th>
                    <th className="border px-2 py-1.5 text-right">Layer</th>
                    <th className="border px-2 py-1.5 text-right">Sonali</th>
                    <th className="border px-2 py-1.5 text-right">Cattle</th>
                    <th className="border px-2 py-1.5 text-right">Fish</th>
                    <th className="border px-2.5 py-1.5 text-right bg-emerald-50">
                      Total Sale
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((c: any) => (
                    <tr key={c.name} className="hover:bg-slate-50">
                      <td className="border px-2.5 py-1.5 font-semibold">
                        {c.name}
                      </td>
                      <td className="border px-2 py-1.5 text-center font-bold">
                        {c.agents}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.broiler.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.layer.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.sonali.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.cattle.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.fish.toLocaleString()}
                      </td>
                      <td className="border px-2.5 py-1.5 text-right font-bold text-emerald-700 bg-emerald-50/50">
                        {c.sale.toLocaleString()}
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