// Shared permission helpers for the farmer app.
// Lives at src/utils/accessControl.ts — pages import it as
//   import { ... } from "../../utils/accessControl";
// Dependency-free: no supabase import, no React. Safe to import
// from any page (Farmers, Visits, DocTracking, Agents, Reports).
//
// BINDING RULES (agreed 2026-11-14):
//  1. One user can be assigned MULTIPLE categories (Poultry, Cattle, Aqua).
//  2. Old data (created_by null): officers of the assigned area CAN
//     edit/delete. Once a record is edited, the save stamps
//     created_by = current user, and after that NOBODY except that
//     creator (or manager / super admin) can edit/delete it.
//  3. Map button: EVERYONE can use it. View is also for everyone.
//     (These helpers only gate edit/delete/add — never view/map.)
//  4. Ownership (officer): can edit/delete ONLY what they added
//     (created_by === own id). Others in the assigned area can only
//     VIEW. Manager / super admin can do everything. Applies to
//     everything: Farmer, Agent, Report, everything.
//  5. Category (officer): Add/Edit/Delete only on their own
//     category's visit page. They can VIEW all pages. They can ADD
//     farmers of ANY category (farmer add is NOT category-gated).
//     Only REPORTS are category-gated. Manager / super admin:
//     everything, everywhere.

export type Category = "Poultry" | "Cattle" | "Aqua";

export const ALL_CATEGORIES: Category[] = ["Poultry", "Cattle", "Aqua"];

/** True for 'manager' or 'super_admin'. */
export function isManagerOrAdmin(
  role: string | null | undefined
): boolean {
  return role === "manager" || role === "super_admin";
}

/**
 * Can this user edit/delete an existing record?
 * - manager / super admin: always true.
 * - 'viewer' (legacy view-only role): NEVER true, even for old
 *   rows with created_by null.
 * - created_by null / undefined / '' (old data, rule 2): any
 *   officer of the assigned area can edit/delete. The save must
 *   then stamp created_by = current user, so the next edit is
 *   creator-only.
 * - otherwise: only the creator (created_by === currentUserId).
 */
export function canEditRecord(
  role: string | null | undefined,
  createdBy: string | null | undefined,
  currentUserId: string | null | undefined
): boolean {
  if (role === "viewer") return false;
  if (isManagerOrAdmin(role)) return true;
  if (!createdBy) return true;
  if (currentUserId && createdBy === currentUserId) return true;
  return false;
}

/**
 * Can this user ADD a report of the given category?
 * - manager / super admin: always true.
 * - 'viewer': never true.
 * - officer: only if the category is in their assigned_categories.
 *   (assignedCategories null/undefined/empty => no report add.)
 */
export function canAddReport(
  role: string | null | undefined,
  assignedCategories: string[] | null | undefined,
  category: string
): boolean {
  if (role === "viewer") return false;
  if (isManagerOrAdmin(role)) return true;
  if (!assignedCategories) return false;
  return assignedCategories.includes(category);
}

/**
 * Can this user ADD a farmer?
 * Farmer add is NOT category-gated (rule 5): any officer,
 * manager or super admin can add a farmer of any category.
 * 'viewer' (and no role) cannot.
 */
export function canAddFarmer(role: string | null | undefined): boolean {
  return role === "officer" || isManagerOrAdmin(role);
}
