import { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import {
  Users,
  UserCheck,
  Building2,
  Calendar,
  MapPin,
  PieChart,
  ChevronDown,
} from "lucide-react";

function parseList(val: any): string[] {
  if (!val) return [];
  if (Array.isArray(val))
    return val.map((s) => String(s).trim()).filter(Boolean);
  if (typeof val === "string") {
    try {
      const parsed = JSON.parse(val);
      if (Array.isArray(parsed)) {
        return parsed.map((s) => String(s).trim()).filter(Boolean);
      }
    } catch {
      return val
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    }
  }
  return [];
}

function formatMonthName(ym: string) {
  const [year, month] = ym.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  return date.toLocaleString("en-US", { month: "short", year: "numeric" });
}

export default function DashboardPage() {
  const [loading, setLoading] = useState(true);

  const now = new Date();
  const runningYm = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(
    2,
    "0"
  )}`;

  // Default selected month is always the Running Month on open/reload
  const [selectedYm, setSelectedYm] = useState<string>(runningYm);
  const [availableMonths, setAvailableMonths] = useState<string[]>([runningYm]);

  // Scoped raw DOC arrays so changing the month dropdown updates stats instantly
  const [scopedDocSupplies, setScopedDocSupplies] = useState<any[]>([]);
  const [scopedDocTrackings, setScopedDocTrackings] = useState<any[]>([]);

  const [baseStats, setBaseStats] = useState({
    totalAgents: 0,
    totalFarmers: 0,
  });

  const [companyStats, setCompanyStats] = useState<any[]>([]);
  const [upazilaStats, setUpazilaStats] = useState<any[]>([]);
  const [monthlyDocStats, setMonthlyDocStats] = useState<any[]>([]);

  useEffect(() => {
    loadDashboard();
  }, []);

  async function loadDashboard() {
    setLoading(true);
    try {
      const {
        data: { user },
      } = await supabase.auth.getUser();

      let role = "officer";
      let uDistricts: string[] = [];
      let uUpazilas: string[] = [];

      if (user) {
        const { data: profile } = await supabase
          .from("users")
          .select("role, districts, upazilas")
          .eq("id", user.id)
          .single();

        if (profile) {
          role = profile.role || "officer";
          uDistricts = parseList(profile.districts);
          uUpazilas = parseList(profile.upazilas);
        }
      }

      const normUpazilas = uUpazilas.map((u) => u.toLowerCase());
      const normDistricts = uDistricts.map((d) => d.toLowerCase());

      const isAllowedArea = (item: any) => {
        const itemUpa = String(item?.upazila || "").trim().toLowerCase();
        const itemDist = String(item?.district || "").trim().toLowerCase();

        if (normUpazilas.length > 0) {
          return normUpazilas.includes(itemUpa);
        }
        if (normDistricts.length > 0) {
          return normDistricts.includes(itemDist);
        }
        return role === "super_admin" || role === "manager";
      };

      const [
        { data: companiesData },
        { data: agentsData },
        { data: farmersData },
        { data: docAgentsData },
        { data: docSuppliesData },
        { data: docTrackingsData },
      ] = await Promise.all([
        supabase.from("companies").select("id, company_name, logo_url"),
        supabase
          .from("agents")
          .select("id, company_id, agent_name, district, upazila"),
        supabase
          .from("farmers")
          .select("id, company_id, district, upazila, farm_category, status"),
        supabase.from("doc_agents").select("doc_id, district, upazila"),
        supabase.from("doc_supplies").select("*"),
        supabase.from("doc_trackings").select("*"),
      ]);

      const companies = companiesData || [];
      const agents = (agentsData || []).filter(isAllowedArea);
      const farmers = (farmersData || []).filter(isAllowedArea);
      const docAgents = (docAgentsData || []).filter(isAllowedArea);

      const allowedDocIds = new Set(
        docAgents.map((da: any) => String(da.doc_id).trim())
      );

      const hasAreaRestriction =
        normUpazilas.length > 0 ||
        normDistricts.length > 0 ||
        (role !== "super_admin" && role !== "manager");

      const docSupplies = hasAreaRestriction
        ? (docSuppliesData || []).filter((s: any) =>
            allowedDocIds.has(String(s.doc_id).trim())
          )
        : docSuppliesData || [];

      const docTrackings = hasAreaRestriction
        ? (docTrackingsData || []).filter((t: any) =>
            allowedDocIds.has(String(t.doc_id).trim())
          )
        : docTrackingsData || [];

      setScopedDocSupplies(docSupplies);
      setScopedDocTrackings(docTrackings);

      setBaseStats({
        totalAgents: agents.length,
        totalFarmers: farmers.length,
      });

      // Collect all months present in database + Running Month
      const monthSet = new Set<string>([runningYm]);
      docSupplies.forEach((s: any) => {
        if (s.h_date) monthSet.add(String(s.h_date).slice(0, 7));
      });
      docTrackings.forEach((t: any) => {
        if (t.h_date) monthSet.add(String(t.h_date).slice(0, 7));
      });

      const sortedMonths = Array.from(monthSet).sort().reverse();
      setAvailableMonths(sortedMonths);

      // 1. Company-wise Agents & Farmers
      const compMap: Record<string, any> = {};
      companies.forEach((c) => {
        compMap[String(c.id)] = {
          id: c.id,
          companyName: c.company_name,
          logoUrl: c.logo_url,
          agents: 0,
          farmers: 0,
          poultry: 0,
          cattle: 0,
          fish: 0,
        };
      });

      agents.forEach((a) => {
        const cid = String(a.company_id);
        if (compMap[cid]) compMap[cid].agents += 1;
      });

      farmers.forEach((f) => {
        const cid = String(f.company_id);
        if (compMap[cid]) {
          compMap[cid].farmers += 1;
          if (f.farm_category === "Poultry") compMap[cid].poultry += 1;
          if (f.farm_category === "Cattle") compMap[cid].cattle += 1;
          if (f.farm_category === "Fish") compMap[cid].fish += 1;
        }
      });

      const compArray = Object.values(compMap)
        .filter((c: any) => c.agents > 0 || c.farmers > 0)
        .sort((a: any, b: any) => b.farmers - a.farmers || b.agents - a.agents);

      setCompanyStats(compArray);

      // 2. District & Upazila-wise Farmers
      const upaMap: Record<string, any> = {};
      const ensureUpa = (district: string, upazila: string) => {
        const key = `${district || "Unknown"}__${upazila || "Unknown"}`;
        if (!upaMap[key]) {
          upaMap[key] = {
            district: district || "-",
            upazila: upazila || "-",
            agents: 0,
            farmers: 0,
          };
        }
        return upaMap[key];
      };

      agents.forEach((a) => {
        if (!a.upazila && !a.district) return;
        ensureUpa(a.district, a.upazila).agents += 1;
      });

      farmers.forEach((f) => {
        if (!f.upazila && !f.district) return;
        ensureUpa(f.district, f.upazila).farmers += 1;
      });

      const upaArray = Object.values(upaMap).sort(
        (a: any, b: any) =>
          b.farmers - a.farmers || a.upazila.localeCompare(b.upazila)
      );
      setUpazilaStats(upaArray);

      // 3. Total DOC by Months
      const monthMap: Record<string, any> = {};
      docSupplies.forEach((s: any) => {
        if (!s.h_date) return;
        const ym = String(s.h_date).slice(0, 7);
        if (!monthMap[ym]) {
          monthMap[ym] = {
            ym,
            label: formatMonthName(ym),
            totalDoc: 0,
            trackedDoc: 0,
          };
        }
        monthMap[ym].totalDoc += Number(s.qty || 0);
      });

      docTrackings.forEach((t: any) => {
        if (!t.h_date) return;
        const ym = String(t.h_date).slice(0, 7);
        if (!monthMap[ym]) {
          monthMap[ym] = {
            ym,
            label: formatMonthName(ym),
            totalDoc: 0,
            trackedDoc: 0,
          };
        }
        monthMap[ym].trackedDoc += Number(t.qty || 0);
      });

      const monthArray = Object.values(monthMap)
        .sort((a: any, b: any) => b.ym.localeCompare(a.ym))
        .map((m: any) => {
          const untracked = Math.max(0, m.totalDoc - m.trackedDoc);
          const pct =
            m.totalDoc > 0
              ? ((m.trackedDoc / m.totalDoc) * 100).toFixed(1)
              : "0.0";
          return { ...m, untracked, pct };
        });

      setMonthlyDocStats(monthArray);
    } finally {
      setLoading(false);
    }
  }

  // Calculate DOC supply & tracked for the currently selected month in dropdown
  const selectedMonthDocSupply = scopedDocSupplies
    .filter((s: any) => String(s.h_date || "").startsWith(selectedYm))
    .reduce((sum: number, s: any) => sum + Number(s.qty || 0), 0);

  const selectedMonthDocTracked = scopedDocTrackings
    .filter((t: any) => String(t.h_date || "").startsWith(selectedYm))
    .reduce((sum: number, t: any) => sum + Number(t.qty || 0), 0);

  const selectedMonthDocPct =
    selectedMonthDocSupply > 0
      ? ((selectedMonthDocTracked / selectedMonthDocSupply) * 100).toFixed(1)
      : "0.0";

  const selectedMonthLabel = formatMonthName(selectedYm);

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-4">
      {/* ==================================================
          INTERACTIVE MONTH SELECTOR DROPDOWN AT TOP RIGHT
      ================================================== */}
      <div className="flex items-center justify-end">
        <div className="relative inline-flex items-center bg-white border border-slate-200 hover:border-emerald-400 rounded-full px-3.5 py-1.5 shadow-sm transition">
          <Calendar size={14} className="text-emerald-600 mr-1.5 shrink-0" />
          <span className="text-xs font-medium text-slate-500 mr-1">
            Month:
          </span>
          <select
            value={selectedYm}
            onChange={(e) => setSelectedYm(e.target.value)}
            className="appearance-none bg-transparent pr-5 text-xs font-bold text-slate-800 outline-none cursor-pointer"
          >
            {availableMonths.map((ym) => (
              <option key={ym} value={ym}>
                {formatMonthName(ym)}
                {ym === runningYm ? " (Running)" : ""}
              </option>
            ))}
          </select>
          <ChevronDown
            size={13}
            className="text-slate-400 pointer-events-none absolute right-3"
          />
        </div>
      </div>

      {/* ==================================================
          1. MINIMAL TOP KPI CARDS (DYNAMIC BY SELECTED MONTH)
      ================================================== */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
        <div className="bg-white border border-slate-200/90 rounded-2xl p-4 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-medium text-slate-500">Total Agents</p>
            <p className="text-2xl sm:text-3xl font-bold text-slate-800 mt-1">
              {baseStats.totalAgents.toLocaleString()}
            </p>
          </div>
          <div className="w-11 h-11 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center shrink-0">
            <Users size={20} />
          </div>
        </div>

        <div className="bg-white border border-slate-200/90 rounded-2xl p-4 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-medium text-slate-500">Total Farmers</p>
            <p className="text-2xl sm:text-3xl font-bold text-slate-800 mt-1">
              {baseStats.totalFarmers.toLocaleString()}
            </p>
          </div>
          <div className="w-11 h-11 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0">
            <UserCheck size={20} />
          </div>
        </div>

        <div className="bg-white border border-slate-200/90 rounded-2xl p-4 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-medium text-slate-500">
              DOC ({selectedMonthLabel})
            </p>
            <p className="text-2xl sm:text-3xl font-bold text-slate-800 mt-1">
              {selectedMonthDocSupply.toLocaleString()}
            </p>
          </div>
          <div className="w-11 h-11 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center shrink-0">
            <Calendar size={20} />
          </div>
        </div>

        <div className="bg-white border border-slate-200/90 rounded-2xl p-4 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-medium text-slate-500">
              Tracked ({selectedMonthLabel})
            </p>
            <p className="text-2xl sm:text-3xl font-bold text-emerald-700 mt-1">
              {selectedMonthDocPct}%
            </p>
            <p className="text-[11px] text-slate-400 font-medium mt-0.5">
              Qty: {selectedMonthDocTracked.toLocaleString()}
            </p>
          </div>
          <div className="w-11 h-11 rounded-xl bg-teal-50 text-teal-600 flex items-center justify-center shrink-0">
            <PieChart size={20} />
          </div>
        </div>
      </div>

      {/* ==================================================
          2. COMPANY-WISE AGENT & FARMER TABLE
      ================================================== */}
      <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100 flex items-center gap-2 bg-slate-50/60">
          <Building2 size={16} className="text-emerald-600" />
          <h2 className="text-sm font-bold text-slate-800">
            Company-wise Agents & Farmers
          </h2>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs sm:text-sm">
            <thead>
              <tr className="border-b border-slate-200 text-slate-500 bg-slate-50/40">
                <th className="px-4 py-2.5 font-semibold">Company</th>
                <th className="px-3 py-2.5 text-center font-semibold">
                  Agents
                </th>
                <th className="px-3 py-2.5 text-center font-semibold">
                  Total Farmers
                </th>
                <th className="px-3 py-2.5 text-center font-semibold">
                  🐔 Poultry
                </th>
                <th className="px-3 py-2.5 text-center font-semibold">
                  🐄 Cattle
                </th>
                <th className="px-3 py-2.5 text-center font-semibold">
                  🐟 Fish
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr>
                  <td colSpan={6} className="py-8 text-center text-slate-400">
                    Loading...
                  </td>
                </tr>
              ) : companyStats.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-8 text-center text-slate-400">
                    No records found in your assigned area.
                  </td>
                </tr>
              ) : (
                companyStats.map((c) => (
                  <tr key={c.id} className="hover:bg-slate-50/80 transition">
                    <td className="px-4 py-2.5 font-semibold text-slate-800">
                      <div className="flex items-center gap-2.5">
                        {c.logoUrl ? (
                          <img
                            src={c.logoUrl}
                            alt={c.companyName}
                            className="w-6 h-6 object-contain rounded border border-slate-200 bg-white p-0.5 shrink-0"
                          />
                        ) : (
                          <div className="w-6 h-6 rounded bg-emerald-50 text-emerald-700 flex items-center justify-center font-bold text-xs shrink-0">
                            {c.companyName?.charAt(0)}
                          </div>
                        )}
                        <span>{c.companyName}</span>
                      </div>
                    </td>
                    <td className="px-3 py-2.5 text-center font-bold text-blue-600">
                      {c.agents}
                    </td>
                    <td className="px-3 py-2.5 text-center font-bold text-emerald-700">
                      {c.farmers}
                    </td>
                    <td className="px-3 py-2.5 text-center text-slate-600">
                      {c.poultry}
                    </td>
                    <td className="px-3 py-2.5 text-center text-slate-600">
                      {c.cattle}
                    </td>
                    <td className="px-3 py-2.5 text-center text-slate-600">
                      {c.fish}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* ==================================================
          3. UPAZILA-WISE FARMERS & 4. TOTAL DOC BY MONTHS
      ================================================== */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 sm:gap-5">
        {/* Upazila-wise Farmers */}
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-100 flex items-center gap-2 bg-slate-50/60">
            <MapPin size={16} className="text-blue-600" />
            <h2 className="text-sm font-bold text-slate-800">
              Assigned Upazila-wise Agents & Farmers
            </h2>
          </div>

          <div className="overflow-x-auto max-h-80">
            <table className="w-full text-left border-collapse text-xs sm:text-sm">
              <thead className="sticky top-0 bg-slate-50 border-b border-slate-200 text-slate-500">
                <tr>
                  <th className="px-4 py-2.5 font-semibold">District</th>
                  <th className="px-4 py-2.5 font-semibold">Upazila</th>
                  <th className="px-3 py-2.5 text-center font-semibold">
                    Agents
                  </th>
                  <th className="px-3 py-2.5 text-center font-semibold">
                    Farmers
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {upazilaStats.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="py-8 text-center text-slate-400">
                      No upazila data found.
                    </td>
                  </tr>
                ) : (
                  upazilaStats.map((u, i) => (
                    <tr key={i} className="hover:bg-slate-50/80 transition">
                      <td className="px-4 py-2.5 text-slate-600">
                        {u.district}
                      </td>
                      <td className="px-4 py-2.5 font-semibold text-slate-800">
                        {u.upazila}
                      </td>
                      <td className="px-3 py-2.5 text-center font-semibold text-blue-600">
                        {u.agents}
                      </td>
                      <td className="px-3 py-2.5 text-center font-bold text-emerald-700">
                        {u.farmers}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* Total DOC by Months */}
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-100 flex items-center gap-2 bg-slate-50/60">
            <Calendar size={16} className="text-amber-600" />
            <h2 className="text-sm font-bold text-slate-800">
              Total DOC by Months
            </h2>
          </div>

          <div className="overflow-x-auto max-h-80">
            <table className="w-full text-left border-collapse text-xs sm:text-sm">
              <thead className="sticky top-0 bg-slate-50 border-b border-slate-200 text-slate-500">
                <tr>
                  <th className="px-4 py-2.5 font-semibold">Month</th>
                  <th className="px-3 py-2.5 text-right font-semibold">
                    Total DOC
                  </th>
                  <th className="px-3 py-2.5 text-right font-semibold">
                    Tracked
                  </th>
                  <th className="px-4 py-2.5 text-right font-semibold">
                    Untracked
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {monthlyDocStats.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="py-8 text-center text-slate-400">
                      No DOC data found for your assigned upazila.
                    </td>
                  </tr>
                ) : (
                  monthlyDocStats.map((m) => {
                    const isSelected = m.ym === selectedYm;
                    return (
                      <tr
                        key={m.ym}
                        onClick={() => setSelectedYm(m.ym)}
                        className={`cursor-pointer transition ${
                          isSelected
                            ? "bg-emerald-50/80 font-semibold"
                            : "hover:bg-slate-50/80"
                        }`}
                      >
                        <td className="px-4 py-2.5 font-semibold text-slate-800">
                          {m.label}
                        </td>
                        <td className="px-3 py-2.5 text-right font-bold text-slate-800">
                          {m.totalDoc.toLocaleString()}
                        </td>
                        <td className="px-3 py-2.5 text-right font-semibold text-emerald-700">
                          {m.trackedDoc.toLocaleString()} ({m.pct}%)
                        </td>
                        <td className="px-4 py-2.5 text-right font-medium text-amber-700">
                          {m.untracked.toLocaleString()}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}