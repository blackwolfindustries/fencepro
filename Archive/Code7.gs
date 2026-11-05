// ============================================================
// FencePro Estimator — Google Apps Script Backend (Code.gs)
// Version: 2.19.10  |  Matches FencePro index.html v3.38.0
// v2.19.10 changes: saveJob() and updateStatus() now keep the Jobs row's existing Contact UID and Customer Name when the caller sends them blank. A value the caller does send still wins (renames and relinks work as before). Nothing else changed.
// v2.19.9 changes: new layoutFixDryRun() and layoutFixApply() editor functions. Make the Jobs tab header row match the code's column order (JOB_HEADERS) and move the older rows into that order. Backs up the Jobs tab first. previewJobsLayout() now stops cleanly once the order matches.
// v2.19.8 changes: new READ-ONLY previewJobsLayout() editor function. Sorts Jobs rows into sheet-order vs code-order and shows what a column-order clean-up would do. Changes no data.
// v2.19.7 changes: new READ-ONLY checkJobsColumnOrder() editor function. Shows where the Jobs tab column order differs from the code (JOB_HEADERS). Changes no data.
// v2.19.6 changes: new repairJobsDryRun() and repairJobsApply() editor functions. Fix only the Jobs rows previewJobsRepair() showed as safe (blank Contact UID, status word in Customer Name, one shifted row). Nothing else changes.
// v2.19.5 changes: new READ-ONLY previewJobsRepair() editor function. Shows what a Jobs/Contacts data repair WOULD change. Changes no data.
// v2.19.4 changes: contact link audit now prints record-level detail (row numbers, would-be links, name mismatches, duplicate-UID rows). Still read-only.
// v2.19.3 changes: new READ-ONLY contact link audit (getContactLinkAudit action and runContactLinkAudit editor function).
//   Reports Estimates/Jobs/Invoices with a blank or unknown Contact UID and duplicate contacts. Changes no data, adds no sheets.
// v2.19.2 changes: version bump only (matches index.html v3.38.0, By ZIP analytics tab: sortable tables + customer detail; no backend changes).
// v2.19.0 changes: new Supplier_Contacts sheet (one row per person: Supplier, Contact Name, Email) with
//   getSupplierContacts / saveSupplierContact / deleteSupplierContact (soft delete). Several contacts per supplier.
//
// v2.18.0 (previous) - Matches index.html v3.34.0
// v2.18.0 changes: Bills_Of_Materials gains 'Estimate ID' and 'Customer Name' columns (auto-added).
//   saveBOM stores the linked estimate (jobId = job number); templates are always saved unlinked.
//
// v2.17.1 changes:
//   - No functional change. index.html v3.33.0 adds drag-and-drop and column
//     sorting to bills of materials and a Template option (saved in the
//     existing Status column). No redeploy is needed for this version.
//
// v2.17.0 changes:
//   - New Vendor_Parts sheet (vendor parts catalog). Every part gets its own
//     unique Part ID (VP-0001...). Part Number must be unique across the whole
//     catalog, case-insensitive. Delete only marks a part inactive.
//   - New Bills_Of_Materials sheet. Lines are saved as a snapshot; the Job ID
//     column is blank until the job link is built.
//   - New actions: getVendorParts, saveVendorPart, bulkSaveVendorParts,
//     deleteVendorPart, getBOMs, saveBOM, deleteBOM. Nothing existing was
//     changed. Redeploy required (update the existing deployment).
//
// v2.16.0 changes:
//   - Customer proposal page and its PDF: estimates set to Credit Card or a
//     0% financing term now lead with the CC / Finance price (line amount and
//     total), with Cash / Check as the small second line. Cash / Check
//     estimates are unchanged. Financing terms used to be treated as cash.
//   - New helpers isCardPay_() and priceRowsHtml_(). Redeploy required.
//
// v2.15.7 changes:
//   - No functional change. Credit Card / financing estimates now show the
//     CC / Finance price as the proposal total in index.html.
//
// v2.15.6 changes:
//   - No functional change. The customer email in index.html shows only the
//     Cash / Check price on Cash / Check estimates.
//
// v2.15.5 changes:
//   - No functional change. Review in Gmail now opens the Drafts folder only.
//
// v2.15.4 changes:
//   - No functional change. index.html fixed the company name in the customer
//     email header and changed the default email subject.
//
// v2.15.3 changes:
//   - No functional change. The customer proposal email in index.html now uses
//     the patriotic navy / red / white layout (no logo image).
//
// v2.15.2 changes:
//   - No functional change. The customer proposal email in index.html is now a
//     clean plain layout (no logo or colors).
//
// v2.15.1 changes:
//   - No functional change. Only the Review in Gmail link in index.html was fixed.
//
// v2.15.0 changes:
//   - New createProposalDraft(data) action. Saves the proposal email (same
//     HTML sendProposalEmail would send) as a Gmail DRAFT in the account the
//     script runs as, and returns the account address plus the draft's message
//     id so the app can open it for review. Nothing is sent. Used by the new
//     Review in Gmail button. sendProposalEmail is unchanged.
//
// v2.14.2 changes:
//   - No functional change. The Proposal button in index.html now offers the
//     estimate picker as a preview only; generateProposalToken is untouched and
//     is still called only by Send Proposal Link.
//
// v2.14.1 changes:
//   - No functional change. Job numbers now run past 999 (3 or 4 digits); this
//     file already accepted any length, so only the version line moved.
//
// v2.13.0 changes:
//   - Manual invoices. New saveManualInvoice(data) action creates or edits an
//     invoice that belongs to no job: its own line items, its own payments
//     and its own sales tax, all stored on the invoice row itself. Numbers
//     start at 5000 and are assigned here under a lock, so two devices can
//     never get the same one. Job ID is the fixed marker 'MANUAL' so it can
//     never match a real job. Nine columns are appended to the END of
//     INVOICE_HEADERS (getOrCreateSheet backfills them on an existing sheet).
//     Money is recomputed here from the items, rate and payments, so the
//     stored total, balance and status can never disagree with each other.
//
// v2.12.0 changes:
//   - New updateEstimateContactInfo(data) action: sets the Email and/or Phone
//     cell on a given list of estimate rows. Used when a contact's email or
//     phone changes (edited on the Contact, or typed on an estimate), so every
//     estimate for that customer carries the same details. The v2.11.2
//     updateEstimateEmails action is kept and now just calls this one.
//
// v2.11.2 changes:
//   - New updateEstimateEmails(data) action: sets the Email cell on a given
//     list of estimate rows. Each estimate stores its own copy of the
//     customer's email, so adding or changing a contact's email never
//     reached the estimates already created for that customer. The front
//     end now sends the matching estimate IDs when a contact is saved.
//
// v2.11.1 changes:
//   - A hosted proposal / PDF with several options (tabs A, B, C or the
//     combined view) printed only the FIRST option's address, in the
//     header. When the options are at different job sites, each option
//     now also shows its own "Job Site" line (street, unit, city, state,
//     zip). When every option shares one address nothing changes. Links
//     already sent pick this up automatically, because the page reads the
//     Estimates sheet each time it is opened.
//
// v2.11.0 changes:
//   - Contacts can now carry additional properties (extra service addresses
//     that roll up to one customer, e.g. an owner of several complexes).
//     One new column, "Properties", appended to the END of CONTACT_HEADERS.
//     It holds a JSON text list of {id, label, addr1, addr2, city, state,
//     zip, county}. The existing saveContact action already carries it, so
//     no new backend action was needed. getOrCreateSheet() backfills the
//     new header on an existing Contacts sheet automatically.
//
// v2.10.1 changes:
//   - Found the REAL, deeper reason a proposal kept showing Black Wolf no
//     matter what company was selected: generateProposalToken()'s
//     appendRow() array had one extra blank placeholder before the company
//     value, which pushed it from column 9 ("Company" — where
//     validateToken() and everything downstream actually reads it) into
//     column 10 ("IP Note") instead. That left the real Company cell
//     blank on every single proposal token this app has ever created, for
//     every company, and getCompanyRecord() treats a blank code as "use
//     BWI" — so every hosted proposal page, PDF, and owner notification
//     email has always shown Black Wolf regardless of what was actually
//     selected. The v3.22.1 frontend fix (send the job's own company
//     instead of the stale activeCompany global) was still correct and
//     necessary, but couldn't have fixed the symptom by itself — the
//     right value was being computed and sent, then silently dropped into
//     the wrong column the moment it reached the sheet.
//   - Also added a company check to the existing-token reuse logic in the
//     same function. Without it, generating "a new" link for the same
//     estimate(s)/mode within the 30-day token window would just hand
//     back whatever token already existed — including its old, wrong
//     company — without ever reaching the part of the function that
//     writes a fresh row. This also means Didz doesn't need any manual
//     sheet cleanup for Job 935 specifically: the next time a link is
//     generated for it, the old token's blank/BWI-defaulting company
//     won't match the new request's Batpig, so a genuinely new, correctly
//     -stamped token gets created automatically.
//
// v2.10.0 changes:
//   - New: cancel a job's install schedule and actually remove it from
//     Google Calendar, not just clear it from FencePro. Until now the
//     Calendar event ID that createCalendarEvent() returns when a job is
//     scheduled was never saved anywhere, so there was no way to find that
//     event again later to delete it. Added a 'Calendar Event ID' column
//     to the Jobs sheet (appended at the end, safe/additive) and threaded
//     it through updateStatus() and getJob() the same way every other
//     schedule field already is.
//   - updateStatus() gained two new, opt-in-only flags so this cancel
//     action can do two things every other caller still can't:
//     `allowStatusDowngrade` lets the incoming status win even when it's a
//     lower rank than what's already saved (normally "never downgrade
//     status" is the rule, which would otherwise silently keep a job
//     marked Scheduled forever), and `clearSchedule` actually blanks the
//     scheduled date/crew/time/notes and the calendar event ID instead of
//     the normal "empty incoming value falls back to what's already
//     there" behavior every other field update relies on. Neither flag
//     does anything unless a caller explicitly sets it, so no existing
//     call site changes behavior.
//   - New action `cancelJobCalendarEvent` — reuses the existing
//     cancelEstimateAppointment(data) function as-is (delete a Calendar
//     event by ID is exactly the same operation either way) rather than
//     writing a second copy of that logic, just registered under a name
//     that reads correctly from the job-scheduling side.
//
// v2.9.1 changes:
//   - Didz asked for the phone number and address on a calendar event
//     (both the job-install scheduling event and the new estimate-visit
//     appointment) to be actual clickable links — tap the phone to dial,
//     tap the address to open it in Maps/Waze — instead of plain text.
//     The Calendar API's description field officially supports HTML, so
//     added _calPhoneLink()/_calMapsLink() helpers that build real
//     <a href="tel:..."> and <a href="https://www.google.com/maps/...">
//     tags, used by both createCalendarEvent() and
//     createEstimateAppointment(). The job-install event never had the
//     customer's phone number in its payload at all before this — added
//     custPhone to the frontend's createCalendarEvent call so there's
//     something to link. A Google Maps search URL was used rather than a
//     Waze-only deep link since it works everywhere (desktop, iOS,
//     Android) and still lets a phone with Waze set as the default maps
//     handler intercept it.
//
// v2.9.0 changes:
//   - New feature: schedule an estimate (site visit) appointment directly
//     from a Contact, before any job/estimate record exists. Didz asked
//     how to tap the existing Google Calendar integration for this rather
//     than build something separate. Added createEstimateAppointment(data)
//     and cancelEstimateAppointment(data) as their own new functions —
//     deliberately NOT reusing createCalendarEvent() itself, since that
//     one is wired to the job-install scheduling flow's date/slot/crew
//     shape and is already working in production; a separate function
//     avoids any chance of regressing it. Both call the same
//     CalendarApp.getDefaultCalendar() the install scheduler already uses,
//     so appointments land on the same calendar Didz is already checking.
//     The appointment (date/time, notes, and the calendar event ID so it
//     can be rescheduled or cancelled later) is stored on the Contact
//     record itself — three columns appended to CONTACT_HEADERS
//     (Est Appt DateTime, Est Appt Notes, Est Appt Event ID). Appending is
//     additive and safe: getOrCreateSheet() already auto-backfills any
//     header missing from an existing sheet, the same mechanism that's
//     been quietly handling every other column addition in this file.
//
// v2.8.3 changes:
//   - Didz asked "what else isn't connected" after v2.8.2 found
//     updateInvoice missing — audited every action name the front end
//     calls against this file's doPost/doGet routing, end to end, rather
//     than spot-checking. Found two more real gaps, both fixed now:
//     deleteInvoice(data) (the Invoices tab delete button was relying on
//     a browser-local tombstone list because the row was never actually
//     removed from the sheet — it would reappear on any other device)
//     and sendCIFormPDF(data) (the "Email ST-124 to Customer" button —
//     this one wasn't silent, it showed a visible "Unknown action" error,
//     so it's a never-finished feature rather than a silent data
//     problem). Every other action name the frontend calls has a real
//     handler — some via doPost's switch, some via doGet's if-chain
//     (getJobs/getInvoices/getJobCosts/pdf) — confirmed by cross-checking
//     the complete list of both sides, not just the one that was reported.
//
// v2.8.2 changes:
//   - New updateInvoice(data) action, and registered in doPost()'s switch.
//     The frontend has been calling this action for a while (from
//     confirmFinalPayment() and the Edit Invoice screen) but it never
//     existed here, so it fell into the default "Unknown action" case —
//     meaning no invoice's Payment Status/Amount Due has ever actually
//     been written back to the Invoices sheet. Updates just the fields
//     sent (Payment Status, Amount Due, Payment Method, Payment Date),
//     found by Invoice ID, and treats 0 as a real value (not "missing")
//     so "Amount Due: 0" (paid in full) is never silently dropped.
//
// v2.8.1 changes:
//   - New updateCompanyForJob(data) action — sets the Company cell on every
//     Estimates row sharing a Job Prefix in one call, so switching a
//     company on any one lettered option (A/B/C...) can apply to every
//     option of that job (Didz: one job never spans multiple companies),
//     without re-running the full saveJob() per sibling.
//
// v2.8.0 changes:
//   - Company selection now flows all the way through, not just to the
//     Estimates sheet's Franchise-style flag. Added 'Company' column to
//     ESTIMATE_HEADERS (saveJob) and 'Company Name'/'Company Address'/
//     'Company Phone' columns to INVOICE_HEADERS (saveInvoice) — the
//     front end already computed these three invoice fields but had no
//     column to store them in, so they were silently dropped on every
//     save.
//   - sendProposalEmail(): the Gmail "From" display name was hardcoded to
//     'Black Wolf Industries LLC' for every company. Now uses the
//     company name passed from the front end.
//   - serveProposalPage(): the hosted/public proposal page a customer
//     actually opens had its own separate hardcoded company logic — a
//     binary TRF-or-else-BWI name check (so Batpig showed as BWI) and a
//     hardcoded BWI/TRF address, in two places, plus the page footer.
//     Now looks up the real company record (name/address/phone) from
//     the same Setup ▸ Manage Companies data (Setup_Data sheet) the
//     front end already uses, via a new getCompanyRecord(code) helper.
//
// v2.7.4 changes:
//   - Analytics: getAnalytics() now formats real Date-typed cells as ISO
//     (yyyy-MM-dd) instead of letting them fall through to JavaScript's
//     verbose default Date-to-string conversion. Sheets auto-converts
//     date-like CSV text (e.g. "8/15/2026") into a real Date cell on
//     write; reading it back with a naive String(cell) produced a format
//     ("Sat Aug 15 2026 00:00:00 GMT-0400...") the front end's date parser
//     couldn't read, so every Analytics date-range filter except "All
//     Time" silently showed zero records on any device other than the one
//     that originally uploaded the CSV. Confirmed via cross-device/
//     cross-browser reports (iPad Safari and desktop Chrome both showed
//     the identical symptom, ruling out a per-device cache issue) plus a
//     full code trace of getDateRange()/parseAnalyticsDate()/renderKPIs()
//     in index.html. Additive fix — only affects how Date cells are
//     serialized on the way out; storage/display in the Sheet itself is
//     unchanged.
//
// v2.7.3 changes:
//   - Job Costs: added a real 'Vendor' column to the Job_Costs sheet
//     (appended at the end — existing rows/columns untouched, per the
//     column-safety rule). saveJobCost()/getJobCosts() now read/write it
//     as its own field instead of it being silently dropped. Previously
//     the front end collected a Vendor value but the sheet had no column
//     for it, so saveJobCost() ignored it entirely and the front end
//     covered the gap by mashing "Vendor — Description" into the
//     Description field and always reporting vendor:'' back on reload.
//
// v2.7.2 changes:
//   - Fixed "Invalid Date" on Deposit Receipts (and any other date field
//     read through getJob()/getJobs()). Sheets auto-converts date-looking
//     strings into real Date cells; reading them back gave a raw JS Date
//     object that serialized to a full timestamp instead of a plain date.
//     Front-end code appending 'T12:00:00' to that timestamp produced an
//     unparseable string. Added a shared sheetVal() helper (same pattern
//     already used in getInvoices()/getJobCosts()/getPayments()) and
//     applied it in getJob() and getJobs() so date-typed cells always
//     come back as plain 'YYYY-MM-DD' strings.
//
// v2.7.1 changes:
//   - Fixed recordProposalResponse() writing a hand-built, out-of-sync
//     positional row to the Jobs sheet on customer portal accept/decline.
//     It was missing a Contact UID slot, which shifted every subsequent
//     column — Status ended up holding a timestamp instead of
//     'Accepted'/'Lost', so accepted proposals never showed as accepted
//     in the app even though the notification email was correct.
//     Now routes through updateStatus() (header-name-safe, same path the
//     in-app status changes use) instead of duplicating row construction.
//
// v2.6 changes:
//   - Added Zip_Lookup sheet (self-growing zip→city/county table)
//   - getZipLookup() wired into getAll()
//   - saveZipEntry() hooked into saveJob() and saveContact()
//   - seedZipLookupFromLegacyTables() — run once manually to
//     pre-populate from the old hardcoded JS tables
//
// DEPLOY AS:
//   Deploy → New deployment → Type: Web App
//   Execute as: Me
//   Who has access: Anyone
// ============================================================

// ── ALLOWED EMAILS — controls who can log into FencePro ──
// Owner is always allowed regardless of the sheet list
const OWNER_EMAIL = 'anthonyd@blackwolfindustriesli.com';

// ── API KEY ──
const API_KEY = 'fp_bwi_2026_xK9mQpLvNzRtWjYd';

// ── YOUR IDs GO HERE ──────────────────────────────────────
const SHEET_ID       = '1Z0Sm5KLmBpo9YQJz13kD2PoQ3sUh4xid24M_4TWG2mg';
const ROOT_FOLDER_ID = '1la_lkMmCftTfN5HepiW-a1uBvpxwLEUC';

// ── CACHED OBJECTS ─────────────────────────────────────────
let _ss         = null;
let _rootFolder = null;
function getSpreadsheet()    { if (!_ss)         _ss         = SpreadsheetApp.openById(SHEET_ID);       return _ss; }
function getRootFolderCached(){ if (!_rootFolder) _rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID); return _rootFolder; }
function getRootFolder()     { return getRootFolderCached(); }

// ── TAB NAMES ──────────────────────────────────────────────
const ESTIMATES_TAB  = 'Estimates';
const JOBS_TAB       = 'Jobs';
const CONTACTS_TAB   = 'Contacts';
const INVOICES_TAB   = 'Invoices';
const SETTINGS_TAB   = 'Settings';
const TOKENS_TAB     = 'Proposal_Tokens';
const JOB_COSTS_TAB  = 'Job_Costs';
const SETUP_DATA_TAB = 'Setup_Data';
const PAYMENTS_TAB   = 'Payments';
const ZIP_LOOKUP_TAB = 'Zip_Lookup';

// ── ZIP LOOKUP HEADERS ─────────────────────────────────────
const ZIP_LOOKUP_HEADERS = ['Zip', 'City', 'County', 'Updated At'];

// ── YOUR NOTIFICATION EMAIL ────────────────────────────────
const NOTIFY_EMAIL = 'anthonyd@blackwolfindustriesli.com';

// ── JOB COSTS HEADERS ──────────────────────────────────────
// 'Vendor' appended at the end (v2.7.3) — column-safety rule: append only,
// never insert, so existing Job_Costs rows/columns are untouched.
const JOB_COST_HEADERS = [
  'Cost ID', 'Job ID', 'Estimate ID',
  'Type', 'Description', 'Date', 'Amount',
  'Created At', 'Updated At', 'Vendor'
];

// ── PAYMENT HEADERS ────────────────────────────────────────────
const PAYMENT_HEADERS = [
  'Payment ID', 'Job ID', 'Invoice ID', 'Type',
  'Amount', 'Method', 'Date', 'Ref',
  'Voided', 'Created At', 'Migrated From'
];

// ── ESTIMATES HEADERS (64 cols) ──
// v2.8.0: 'Company' appended at the end (never inserted — see Migration
// safety standing rule) — the company code (BWI/TRF/BATPIG/etc.) this
// estimate belongs to. Blank on any pre-v2.8.0 row; front end treats a
// blank/missing value as 'BWI', matching prior behavior exactly.
const ESTIMATE_HEADERS = [
  'Estimate ID', 'Job Prefix', 'Job No', 'Contact UID', 'Saved At',
  'Customer Name', 'Phone', 'Email', 'Job Address', 'Job Address 2',
  'Job City', 'Job State', 'Zip', 'County',
  'Job Type', 'Fence Type', 'Payment Method', 'Supplier',
  'Total LF', 'Num Gates', 'Markup Pct', 'Franchise On',
  'Material Cost', 'Labor Cost', 'Demo LF', 'Demo Rate',
  'Concrete Bags', 'Carting Fee', 'AddOn Costs', 'Discount',
  'Supplier CC Rate', 'Delivery Fee', 'Mat Del PreTax',
  'Material Tax', 'Supplier Fee', 'Concrete Total',
  'Base SM Price', 'Effective Markup Pct', 'Final Price CC',
  'Cash Price', 'Customer Tax', 'Final With Tax',
  'Gross Margin Pct', 'Gross Margin Dollars',
  'Material Pct', 'Labor Pct',
  'Scope Notes', 'Project Description', 'Delivery Override',
  'Channel', 'Campaign',
  'Photos Count', 'Drive Folder', 'Internal Notes', 'Photo URLs', 'Doc URLs',
  'Job ID',
  'Status',
  'Deposit Type',
  'Deposit Required',
  'Contract Price',
  'Contract Price Type',
  'Fence-Only Price Per LF',
  'Material',
  'Company',
  'Created At'   // v2.14.0: when the estimate row was first created; never overwritten on update
];

// ── JOBS HEADERS (23 cols) ──
const JOB_HEADERS = [
  'Job ID', 'Estimate ID', 'Contact UID', 'Customer Name',
  'Status', 'Status Changed At', 'Changed By',
  'Contract Price', 'Contract Price Type', 'Deposit Type', 'Deposit Required',
  'Deposit Paid', 'Deposit Method', 'Deposit Date', 'Deposit Ref',
  'Scheduled Date', 'Scheduled Crew', 'Scheduled Time', 'Schedule Notes',
  'Final Payment Amount', 'Final Payment Method', 'Final Payment Date', 'Final Payment Ref',
  'Balance Due', 'Lost Reason', 'Notes',
  'Calendar Event ID' // v2.10.0 — the Google Calendar event ID for this job's
                       // install appointment, so it can actually be found and
                       // deleted later when Didz cancels a scheduled job.
];

// ── INVOICE HEADERS ────────────────────────────────────────
// v2.8.0: 'Company Name'/'Company Address'/'Company Phone' appended —
// the front end already computed these three fields at invoice-creation
// time (from the estimate's selected company) but had no column to
// store them in, so they were silently dropped on every save and every
// invoice fell back to a single global default. Storing the resolved
// strings (not just a company code) so a historical invoice stays
// correct even if that company's address/phone changes later in Setup.
const INVOICE_HEADERS = [
  'Invoice ID', 'Invoice Number', 'Job ID', 'Job Prefix', 'Contact UID',
  'Customer Name', 'Email', 'Job Address', 'Job City', 'Job State', 'Job Zip',
  'Invoice Date', 'Due Date',
  'Fence Type', 'Final Price CC', 'Cash Price', 'Customer Tax', 'Amount Due',
  'Payment Status', 'Payment Date', 'Payment Method', 'Payment Instructions',
  'Customer Note', 'Drive Folder',
  'Company Name', 'Company Address', 'Company Phone',
  // v2.13.0 - manual invoices (appended; never inserted)
  'Invoice Type', 'Line Items', 'Payments', 'Subtotal', 'Tax County', 'Tax Rate',
  'Invoice Total', 'Customer Phone', 'Job Address 2', 'Company'
];

// ── CONTACT HEADERS ────────────────────────────────────────
const CONTACT_HEADERS = [
  'UID', 'Name', 'Phone', 'Alt Phone', 'Email', 'Company', 'Category', 'Nickname',
  'Address 1', 'Address 2', 'City', 'State', 'Zip', 'County', 'Community',
  'Bill Address 1', 'Bill Address 2', 'Bill City', 'Bill State', 'Bill Zip',
  'Channel', 'Campaign', 'Referred By', 'Tags',
  'Contact Type', 'Project Type', 'Projected LF', 'Fence Kind', 'Fence Style',
  'Services', 'Gate Code',
  'Customer Since', 'Last Service Date', 'Lifetime Value', 'Balance Due',
  'Tax Rate', 'Card On File', 'Auto Pay', 'Do Not Email', 'Do Not Text',
  'Flash Message', 'Additional Contacts',
  'Created At', 'Source', 'Inactive',
  'Est Appt DateTime', 'Est Appt Notes', 'Est Appt Event ID',
  'Properties'
];

// ══════════════════════════════════════════════════════════════
// HTTP ROUTER
// ══════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e?.parameter?.action || '';
  const token  = e?.parameter?.token  || '';

  if (token && e?.parameter?.action === 'pdf') return generateProposalPdf(token);
  if (token) return serveProposalPage(token);

  const key = e?.parameter?.key || '';
  if (key !== API_KEY) {
    return jsonResponse({ error: 'Unauthorized' });
  }

  if (action === 'getAll')       return jsonResponse(getAll());
  if (action === 'getJobs')      return jsonResponse(getJobs());
  if (action === 'getContacts')  return jsonResponse(getContacts());
  if (action === 'getInvoices')  return jsonResponse(getInvoices());
  if (action === 'getJobCosts')  return jsonResponse(getJobCosts());
  return jsonResponse({ status: 'FencePro Backend v3.1 — OK' });
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'recordProposalResponse') {
      return jsonResponse(recordProposalResponse(body.data));
    }

    if (body.key !== API_KEY) {
      return jsonResponse({ error: 'Unauthorized' });
    }

    let result;
    switch(action) {
      case 'saveJob':              result = saveJob(body.data);                    break;
      case 'updateStatus':         result = updateStatus(body.data);               break;
      case 'getJob':               result = getJob(body.data);                     break;
      case 'deleteJob':            result = deleteJob(body.data);                  break;
      case 'saveInvoice':          result = saveInvoice(body.data);                break;
      case 'updateInvoice':        result = updateInvoice(body.data);              break;
      case 'deleteInvoice':        result = deleteInvoice(body.data);              break;
      case 'saveManualInvoice':    result = saveManualInvoice(body.data);          break;
      case 'sendCIFormPDF':        result = sendCIFormPDF(body.data);              break;
      case 'uploadPhoto':          result = uploadPhoto(body.data);                break;
      case 'uploadDoc':            result = uploadDoc(body.data);                  break;
      case 'saveContact':          result = saveContact(body.data);                break;
      case 'bulkSaveContacts':     result = bulkSaveContacts(body.data);           break;
      case 'deleteContact':        result = deleteContact(body.data);              break;
      case 'getSheetContacts':     result = getSheetContacts();                    break;
      case 'getContactLinkAudit':  result = getContactLinkAudit();                 break;
      case 'saveSettings':         result = saveSettings(body.data);               break;
      case 'createProposalDraft':  result = createProposalDraft(body.data);        break;
      case 'getAll':              result = getAll();                              break;
      case 'getSettings':          result = getSettings();                         break;
      case 'verifyAccess':         result = verifyAccess(body.data);               break;
      case 'saveAllowedEmails':    result = saveAllowedEmailsToSheet(body.data);   break;
      case 'getAllowedEmails':      result = getAllowedEmailsFromSheet();            break;
      case 'removeAllowedEmail':   result = removeAllowedEmailFromSheet(body.data); break;
      case 'generateProposalToken':result = generateProposalToken(body.data);      break;
      case 'recordProposalResponse':result = recordProposalResponse(body.data);    break;
      case 'sendProposalEmail':     result = sendProposalEmail(body.data);         break;
      case 'getProposalAudit':      result = getProposalAudit(body.data);          break;
      case 'saveJobCost':           result = saveJobCost(body.data);               break;
      case 'deleteJobCost':         result = deleteJobCost(body.data);             break;
      case 'getJobCosts':           result = getJobCosts();                        break;
      case 'getPayments':           result = getPayments(body.data);               break;
      case 'voidPayment':           result = voidPayment(body.data);               break;
      case 'getImageAsBase64':      result = getImageAsBase64(body.data);          break;
      case 'saveAnalytics':         result = saveAnalytics(body.data);             break;
      case 'getAnalytics':          result = getAnalytics();                       break;
      case 'createCalendarEvent':   result = createCalendarEvent(body.data);       break;
      case 'createEstimateAppointment': result = createEstimateAppointment(body.data); break;
      case 'cancelEstimateAppointment': result = cancelEstimateAppointment(body.data); break;
      // v2.10.0 — cancelling a scheduled job's calendar event is the exact
      // same operation (delete a Calendar event by ID) as cancelling an
      // estimate appointment, so it reuses that function under a name that
      // reads correctly from the job-scheduling side rather than copying
      // the same three lines of logic a second time.
      case 'cancelJobCalendarEvent':    result = cancelEstimateAppointment(body.data); break;
      // v2.17.0 - vendor parts catalog and bills of materials
      case 'getVendorParts':        result = getVendorParts();                     break;
      case 'saveVendorPart':        result = saveVendorPart(body.data);            break;
      case 'bulkSaveVendorParts':   result = bulkSaveVendorParts(body.data);       break;
      case 'deleteVendorPart':      result = deleteVendorPart(body.data);          break;
      case 'getSupplierContacts':   result = getSupplierContacts();                break;
      case 'saveSupplierContact':   result = saveSupplierContact(body.data);       break;
      case 'deleteSupplierContact': result = deleteSupplierContact(body.data);     break;
      case 'getBOMs':               result = getBOMs();                            break;
      case 'saveBOM':               result = saveBOM(body.data);                   break;
      case 'deleteBOM':             result = deleteBOM(body.data);                 break;
      case 'getJobFiles':           result = getJobFiles(body.data);               break;
      case 'saveSetupData':         result = saveSetupData(body.data);             break;
      case 'getSetupData':          result = getSetupData();                       break;
      case 'updateCompanyForJob':   result = updateCompanyForJob(body.data);       break;
      case 'updateEstimateEmails':  result = updateEstimateEmails(body.data);      break;
      case 'updateEstimateContactInfo': result = updateEstimateContactInfo(body.data); break;
      default:                     result = { error: 'Unknown action: ' + action };
    }
    return jsonResponse(result);
  } catch(err) {
    return jsonResponse({ error: err.toString(), stack: err.stack });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Sheets auto-detects date-looking strings (e.g. "2026-08-13") and stores the
// cell as a real Date value. Reading it back gives a JS Date object, which
// JSON.stringify serializes as a full timestamp ("2026-08-13T04:00:00.000Z").
// Front-end code that expects a plain date and appends 'T12:00:00' to it
// (deposit receipts, schedule display, etc.) then builds an invalid string —
// this is the "Invalid Date" bug. Same guard already used in getInvoices(),
// getJobCosts() and getPayments() — apply here too whenever reading a sheet
// row into an object, so date-typed cells always come back as plain strings.
function sheetVal(v) {
  return (v instanceof Date && !isNaN(v.getTime())) ? v.toISOString().split('T')[0] : v;
}

// ══════════════════════════════════════════════════════════════
// PAYMENTS — write / read / void
// ══════════════════════════════════════════════════════════════

// Internal helper — appends one payment row to the Payments sheet.
// Never call this directly from the router; use updateStatus which
// calls this after writing the Jobs row.
function writePayment(jobId, type, amount, method, date, ref, invoiceId) {
  if (!jobId || !amount || amount <= 0) return null;
  try {
    const sheet   = getOrCreateSheet(PAYMENTS_TAB, PAYMENT_HEADERS);
    const year    = new Date().getFullYear();
    const lastRow = sheet.getLastRow();
    const seqNum  = String(Math.max(0, lastRow - 1) + 1).padStart(4, '0');
    const payId   = 'PAY-' + year + '-' + seqNum;
    const now     = new Date().toISOString();

    // Normalize date — accept ISO string, Date object, or plain date string
    let dateStr = '';
    if (date instanceof Date) {
      dateStr = isNaN(date.getTime()) ? '' : date.toISOString().split('T')[0];
    } else if (date) {
      // If it looks like a full ISO timestamp, trim to date only
      dateStr = String(date).trim().split('T')[0];
    }

    const row = [
      payId,
      String(jobId),
      invoiceId || '',
      type,           // deposit / final / partial / adjustment
      Number(amount),
      method || '',
      dateStr,
      ref    || '',
      '',             // Voided — blank = active
      now,
      ''              // Migrated From — blank for live entries
    ];
    sheet.appendRow(row);
    return payId;
  } catch(e) {
    Logger.log('writePayment error: ' + e.toString());
    return null;
  }
}

// Returns all non-voided payment rows for a given jobId (or all jobs if
// jobId is omitted). Result is an object keyed by jobId for easy lookup.
// Each job entry: { payments: [...], totalPaid: n, depositPaid: n, finalPaid: n }
function getPayments(data) {
  try {
    const sheet = getOrCreateSheet(PAYMENTS_TAB, PAYMENT_HEADERS);
    const raw   = sheet.getDataRange().getValues();
    if (raw.length < 2) return { paymentMap: {}, payments: [] };

    const headers = raw[0].map(h => String(h).trim());
    const filterJobId = data && data.jobId ? String(data.jobId).trim() : null;

    const paymentMap = {};
    const allPayments = [];

    raw.slice(1).forEach(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = row[i];
        if (v instanceof Date) v = isNaN(v.getTime()) ? '' : v.toISOString().split('T')[0];
        obj[h] = v;
      });

      // Skip voided rows
      if (String(obj['Voided'] || '').trim().toLowerCase() === 'yes') return;
      // Skip blank rows
      if (!obj['Payment ID'] && !obj['Job ID']) return;

      const jId  = String(obj['Job ID'] || '').trim();
      if (!jId) return;
      if (filterJobId && jId !== filterJobId) return;

      const entry = {
        paymentId: String(obj['Payment ID'] || ''),
        jobId:     jId,
        invoiceId: String(obj['Invoice ID'] || ''),
        type:      String(obj['Type']   || ''),
        amount:    Number(obj['Amount'] || 0),
        method:    String(obj['Method'] || ''),
        date:      String(obj['Date']   || ''),
        ref:       String(obj['Ref']    || ''),
        createdAt: String(obj['Created At'] || '')
      };

      if (!paymentMap[jId]) {
        paymentMap[jId] = { payments: [], totalPaid: 0, depositPaid: 0, finalPaid: 0 };
      }
      paymentMap[jId].payments.push(entry);
      paymentMap[jId].totalPaid    += entry.amount;
      if (entry.type === 'deposit') paymentMap[jId].depositPaid += entry.amount;
      if (entry.type === 'final' || entry.type === 'partial') paymentMap[jId].finalPaid += entry.amount;
      allPayments.push(entry);
    });

    return { paymentMap, payments: allPayments, count: allPayments.length };
  } catch(e) {
    return { paymentMap: {}, payments: [], error: e.toString() };
  }
}

// Marks a payment row as Voided. Never deletes — audit trail preserved.
function voidPayment(data) {
  try {
    const payId = String(data.paymentId || '').trim();
    if (!payId) return { error: 'No paymentId provided' };
    const sheet = getOrCreateSheet(PAYMENTS_TAB, PAYMENT_HEADERS);
    const row   = findRow(sheet, payId, 0);
    if (row < 0) return { error: 'Payment not found: ' + payId };
    const voidedCol = PAYMENT_HEADERS.indexOf('Voided') + 1; // 1-based for getRange
    sheet.getRange(row, voidedCol).setValue('Yes');
    return { success: true, voided: payId };
  } catch(e) {
    return { error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// SAVE JOB
// ══════════════════════════════════════════════════════════════
function saveJob(data) {
  const estSheet = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const jobSheet = getOrCreateSheet(JOBS_TAB,      JOB_HEADERS);

  const prefix    = (data.jobPrefix || '').trim();
  const year      = new Date().getFullYear();
  const custFirst = (data.custName || 'Unknown').split(' ')[0].replace(/[^a-zA-Z0-9]/g,'');

  const isNew      = !data.id || data.id.trim() === '';
  let existingRow  = isNew ? -1 : findRow(estSheet, data.id, 0);
  let assignedLetter = 'A';
  let finalId      = data.id;

  if (isNew || existingRow === -1) {
    assignedLetter = getNextEstimateLetter(estSheet, prefix);
    finalId        = prefix + '-' + custFirst + '-' + assignedLetter + '-' + year;
    existingRow    = -1;
  } else {
    const parts    = data.id.split('-');
    assignedLetter = parts[2] || 'A';
    finalId        = data.id;
  }

  const jobId = String(finalId).split('-')[0];
  const folderUrl = ensureJobFolder(finalId, data.custName, existingRow === -1);

  // New/unknown zip entered on Estimator — persist it for future lookups.
  // Dedup-safe; no-ops if the zip is already known.
  saveZipEntry(data.custZip, data.custCity || data.jobCity, data.custCounty);

  // Read existing Jobs row BEFORE building estRow so existingCP/existingCPT are available
  const STATUS_RANK_GS = {
    'Draft':0,'Committed':1,'Sent':2,'Accepted':3,
    'Deposit Received':4,'Scheduled':5,'In Progress':6,
    'Ready for Invoice':7,'Invoiced':8,'Paid in Full':9,'Lost':2
  };
  const existingJobRow = findRow(jobSheet, jobId, 0);
  let existingCP=0, existingCPT='', existingDP=0, existingDM='', existingDD='', existingDR='';
  let existingUID='', existingName='';   // v2.19.10 — kept when the caller sends blanks
  if (existingJobRow > 0) {
    const ex = jobSheet.getRange(existingJobRow, 1, 1, JOB_HEADERS.length).getValues()[0];
    existingCP  = Number(ex[JOB_HEADERS.indexOf('Contract Price')])      || 0;
    existingCPT = String(ex[JOB_HEADERS.indexOf('Contract Price Type')] || '');
    existingDP  = Number(ex[JOB_HEADERS.indexOf('Deposit Paid')]        || 0);
    existingDM  = String(ex[JOB_HEADERS.indexOf('Deposit Method')]      || '');
    existingDD  = String(ex[JOB_HEADERS.indexOf('Deposit Date')]        || '');
    existingDR  = String(ex[JOB_HEADERS.indexOf('Deposit Ref')]         || '');
    existingUID  = String(ex[JOB_HEADERS.indexOf('Contact UID')]   || '').trim();   // v2.19.10
    existingName = String(ex[JOB_HEADERS.indexOf('Customer Name')] || '').trim();   // v2.19.10
  }

  // v2.14.0: Created At is stamped once. On update keep what the sheet already
  // has (old rows stay blank rather than get a made-up date).
  const _createdCol = ESTIMATE_HEADERS.indexOf('Created At');
  let createdAtVal = new Date().toISOString();
  if (existingRow > 0) {
    const prevCreated = estSheet.getRange(existingRow, _createdCol + 1).getValue();
    createdAtVal = prevCreated instanceof Date ? prevCreated.toISOString() : String(prevCreated || '');
  }

  const estRow = [
    finalId, data.jobPrefix||'', jobId, data.contactUID||'',
    data.savedAt||new Date().toISOString(),
    data.custName||'', data.custPhone||'', data.custEmail||'',
    data.jobAddr||'', data.jobAddr2||'', data.jobCity||'', data.jobState||'NY',
    data.custZip||'', data.custCounty||'',
    data.jobType||'', data.fenceType||'', data.payMethod||'', data.supplier||'',
    Number(data.totalLF)||0, Number(data.numGates)||0, Number(data.markupPct)||40,
    data.franchiseOn?'Yes':'No',
    Number(data.matCost)||0, Number(data.laborCost)||0,
    Number(data.demoLF)||0, Number(data.demoRate)||3,
    Number(data.concBags)||0, Number(data.cartFee)||0,
    Number(data.addOns)||0, Number(data.discount)||0,
    Number(data.calc_suppCCRate)||0, Number(data.calc_delivery)||0,
    Number(data.calc_matDel)||0, Number(data.calc_matTax)||0,
    Number(data.calc_suppFee)||0, Number(data.calc_concrete)||0,
    Number(data.calc_base)||0, Number(data.calc_effMarkup)||0,
    Number(data.calc_final)||0, Number(data.calc_cash)||0,
    Number(data.calc_custTax)||0, Number(data.calc_finalTax)||0,
    Number(data.calc_gmPct)||0, Number(data.calc_gmDol)||0,
    data.calc_final>0 ? Math.round(data.calc_matTotal/data.calc_final*1000)/10 : 0,
    data.calc_final>0 ? Math.round((data.laborCost+(data.calc_demoDol||0))/data.calc_final*1000)/10 : 0,
    data.scopeNotes||'', data.projDesc||'',
    data.deliveryOverride?'Yes':'No',
    data.custChannel||'', data.custCampaign||'',
    (data.photos||[]).length, folderUrl, data.internalNotes||'',
    (data.photoUrls||[]).join(','),
    (data.docs||[]).filter(d=>d.driveUrl).map(d=>d.driveUrl).join(','),
    jobId,
    (data.propStatus||'Draft').trim(),
    data.depositType          || '50pct',
    Number(data.depositRequired) || 0,
    // Contract Price and Type — preserved from Jobs row if incoming is 0/blank
    // so the accepted price travels with the estimate record permanently
    Number(data.contractPrice)     > 0 ? Number(data.contractPrice)    : existingCP,
    data.contractPriceType              ? data.contractPriceType        : existingCPT,
    Number(data.calc_plfFenceOnly) || 0,
    data.material || '',
    data.company  || 'BWI',
    createdAtVal
  ];

  if (existingRow > 0) {
    estSheet.getRange(existingRow, 1, 1, estRow.length).setValues([estRow]);
  } else {
    estSheet.appendRow(estRow);
  }

  // (existingCP/existingDP read above before estRow)

  const inCP  = Number(data.contractPrice)  || 0;
  const inCPT = data.contractPriceType      || '';
  const inDP  = Number(data.depositPaid)    || 0;

  const jobRow = [
    jobId, finalId||'', data.contactUID||existingUID, data.custName||existingName,
    (data.propStatus||'Draft').trim(),
    new Date().toISOString(), 'FencePro',
    inCP  > 0 ? inCP  : existingCP,
    inCPT      ? inCPT : existingCPT,
    data.depositType || '50pct',
    // If depositType is 'none', always write 0 regardless of what client sends
    data.depositType === 'none' ? 0 : (Number(data.depositRequired) || 0),
    inDP  > 0 ? inDP  : existingDP,
    data.depositMethod  || existingDM,
    data.depositDate    || existingDD,
    data.depositRef     || existingDR,
    data.scheduledDate||'', data.scheduledCrew||'', data.scheduledTime||'', data.scheduleNotes||'',
    Number(data.finalPaymentAmount)||0, data.finalPaymentMethod||'',
    data.finalPaymentDate||'', data.finalPaymentRef||'',
    Number(data.balanceDue)||0, data.lostReason||'', data.notes||''
  ];

  if (existingJobRow > 0) {
    const existingJobData = jobSheet.getRange(existingJobRow, 1, 1, JOB_HEADERS.length).getValues()[0];
    const existingStatus  = String(existingJobData[JOB_HEADERS.indexOf('Status')]||'').trim();
    const incomingStatus  = jobRow[4];
    const existingRank    = STATUS_RANK_GS[existingStatus] !== undefined ? STATUS_RANK_GS[existingStatus] : -1;
    const incomingRank    = STATUS_RANK_GS[incomingStatus] !== undefined ? STATUS_RANK_GS[incomingStatus] : -1;
    if (existingRank > incomingRank) jobRow[4] = existingStatus;
    jobSheet.getRange(existingJobRow, 1, 1, jobRow.length).setValues([jobRow]);
  } else {
    jobSheet.appendRow(jobRow);
  }

  return { success: true, jobId, estimateId: finalId, assignedLetter, folderUrl };
}

// ══════════════════════════════════════════════════════════════
// GET JOBS  — Jobs tab workflow fields WIN over Estimates tab
// for contract/deposit data (these fields only live in Jobs tab)
// ══════════════════════════════════════════════════════════════
function getJobs() {
  const estSheet = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const jobSheet = getOrCreateSheet(JOBS_TAB,      JOB_HEADERS);

  const estData = estSheet.getDataRange().getValues();
  const jobData = jobSheet.getDataRange().getValues();
  if (estData.length <= 1) return { jobs: [], count: 0 };

  const estHeaders = estData[0].map(h => String(h).trim());
  const jobHeaders = jobData.length > 1 ? jobData[0].map(h => String(h).trim()) : [];

  const jobMap = {};
  if (jobData.length > 1) {
    jobData.slice(1).forEach(row => {
      if (!row[0]) return;
      const jId = String(row[0]).trim();
      const obj  = {};
      jobHeaders.forEach((h, i) => { obj[h] = sheetVal(row[i]); });
      jobMap[jId] = obj;
    });
  }

  // Load Payments sheet once and build a map keyed by Job ID
  const paymentsResult = getPayments({});
  const paymentMap     = paymentsResult.paymentMap || {};

  const jobs = estData.slice(1)
    .filter(row => row[0])
    .map(row => {
      const est = {};
      estHeaders.forEach((h, i) => { est[h] = sheetVal(row[i]); });
      // Always use numeric prefix for jobMap lookup — the Jobs sheet key (col A)
      // is always the numeric prefix (e.g. "937"), never the full estimate ID.
      // est['Job ID'] might be "937" or "937-Frank-C-2026" depending on when it was saved.
      const rawJId  = String(est['Job ID']||'').trim() || String(est['Estimate ID']||'').split('-')[0].trim();
      const jId     = rawJId.split('-')[0].trim(); // always use numeric prefix
      const job     = jId ? (jobMap[jId] || {}) : {};

      const merged = { ...job, ...est };
      merged['Status'] = String(est['Status']||'').trim() || String(job['Status']||'').trim() || 'Draft';

      // Jobs tab WINS for workflow fields — these are never written to Estimates sheet.
      // Without this, a portal acceptance (which only writes to Jobs tab) gets zeroed
      // out because the Estimates row has empty values that overwrite via spread above.
      // For Contract Price and Type: prefer whichever source has a non-zero/non-blank value
      // Both Jobs and Estimates now store these — use the best available
      const cpJob  = Number(job['Contract Price'])  || 0;
      const cpEst  = Number(est['Contract Price'])  || 0;
      const cptJob = String(job['Contract Price Type'] || '');
      const cptEst = String(est['Contract Price Type'] || '');
      merged['Contract Price']      = cpJob  > 0 ? cpJob  : (cpEst  > 0 ? cpEst  : 0);
      merged['Contract Price Type'] = cptJob      ? cptJob : (cptEst      ? cptEst : '');

      // Other workflow fields: Jobs tab wins over Estimates tab
      const JOBS_WIN_FIELDS = [
        'Deposit Type', 'Deposit Required',
        'Deposit Paid', 'Deposit Method', 'Deposit Date', 'Deposit Ref',
        'Balance Due', 'Scheduled Date', 'Scheduled Crew', 'Scheduled Time', 'Schedule Notes',
        'Final Payment Amount', 'Final Payment Method', 'Final Payment Date', 'Final Payment Ref',
        'Lost Reason'
      ];
      JOBS_WIN_FIELDS.forEach(field => {
        const jobVal = job[field];
        const estVal = est[field];
        const jobHasData = jobVal !== undefined && jobVal !== null && jobVal !== '' && jobVal !== 0;
        if (jobHasData) merged[field] = jobVal;
        else if (estVal !== undefined) merged[field] = estVal;
      });

      // Attach payments data from Payments sheet
      // payments[] is the full history; totalPaid / depositPaid / finalPaid are convenience sums
      // balanceDue_calc is calculated from Payments sheet — more accurate than stored Balance Due
      const pmtData = jId ? (paymentMap[jId] || null) : null;
      if (pmtData) {
        merged.payments      = pmtData.payments;
        merged.totalPaid     = pmtData.totalPaid;
        merged.depositPaid   = pmtData.depositPaid;   // overrides Jobs sheet value if Payments sheet has data
        merged.finalPaid     = pmtData.finalPaid;
        const cp             = merged['Contract Price'] || 0;
        merged.balanceDue_calc = cp > 0 ? Math.max(0, cp - pmtData.totalPaid) : 0;
      } else {
        merged.payments        = [];
        merged.totalPaid       = Number(merged['Deposit Paid'] || 0) + Number(merged['Final Payment Amount'] || 0);
        merged.depositPaid     = Number(merged['Deposit Paid'] || 0);
        merged.finalPaid       = Number(merged['Final Payment Amount'] || 0);
        merged.balanceDue_calc = merged['Balance Due'] || 0;
      }

      return merged;
    })
    .reverse();

  return { jobs, count: jobs.length };
}

// ══════════════════════════════════════════════════════════════
// UPDATE STATUS
// ══════════════════════════════════════════════════════════════
function updateStatus(data) {
  const sheet = getOrCreateSheet(JOBS_TAB, JOB_HEADERS);

  // STATUS RANK GUARD — never downgrade status, but always update workflow fields
  const STATUS_RANK_US = {
    'Draft':0,'Committed':1,'Sent':2,'Accepted':3,
    'Deposit Received':4,'Scheduled':5,'In Progress':6,
    'Ready for Invoice':7,'Invoiced':8,'Paid in Full':9,'Lost':2
  };
  let finalStatus    = (data.status||'').trim();
  let existingCP     = 0;
  let existingCPT    = '';
  let existingDP     = 0;
  let existingDM     = '';
  let existingDD     = '';
  let existingDR     = '';
  let existingSD     = '';   // Scheduled Date
  let existingSC     = '';   // Scheduled Crew
  let existingST     = '';   // Scheduled Time
  let existingSN     = '';   // Schedule Notes
  let existingFPA    = 0;    // Final Payment Amount
  let existingFPM    = '';   // Final Payment Method
  let existingFPD    = '';   // Final Payment Date
  let existingFPR    = '';   // Final Payment Ref
  let existingBD     = 0;    // Balance Due
  let existingLR     = '';   // Lost Reason
  let existingNotes  = '';   // Notes
  let existingCEI    = '';   // Calendar Event ID
  let existingUID    = '';   // Contact UID   (v2.19.10 — kept when the caller sends blanks)
  let existingName   = '';   // Customer Name (v2.19.10)

  const existingJobRow = findRow(sheet, data.jobId, 0);
  if (existingJobRow > 0) {
    const ex = sheet.getRange(existingJobRow, 1, 1, JOB_HEADERS.length).getValues()[0];
    const existingStatus = String(ex[JOB_HEADERS.indexOf('Status')]||'').trim();
    const existingRank   = STATUS_RANK_US[existingStatus] ?? -1;
    const incomingRank   = STATUS_RANK_US[finalStatus]    ?? -1;
    // Keep higher status — unless this call explicitly asks to allow a
    // deliberate downgrade (v2.10.0: cancelling a job's schedule moves it
    // back a step, e.g. Scheduled -> Deposit Received, which the normal
    // "never downgrade" guard would otherwise silently block).
    if (existingRank > incomingRank && !data.allowStatusDowngrade) finalStatus = existingStatus;
    // Preserve ALL existing workflow fields so a partial update (e.g. portal acceptance
    // that only sends status + contract price) never wipes previously saved schedule or
    // final payment data.
    existingCP   = Number(ex[JOB_HEADERS.indexOf('Contract Price')])       || 0;
    existingCPT  = String(ex[JOB_HEADERS.indexOf('Contract Price Type')]   || '');
    existingDP   = Number(ex[JOB_HEADERS.indexOf('Deposit Paid')]          || 0);
    existingDM   = String(ex[JOB_HEADERS.indexOf('Deposit Method')]        || '');
    existingDD   = String(ex[JOB_HEADERS.indexOf('Deposit Date')]          || '');
    existingDR   = String(ex[JOB_HEADERS.indexOf('Deposit Ref')]           || '');
    existingSD   = String(ex[JOB_HEADERS.indexOf('Scheduled Date')]        || '');
    existingSC   = String(ex[JOB_HEADERS.indexOf('Scheduled Crew')]        || '');
    existingST   = String(ex[JOB_HEADERS.indexOf('Scheduled Time')]        || '');
    existingSN   = String(ex[JOB_HEADERS.indexOf('Schedule Notes')]        || '');
    existingFPA  = Number(ex[JOB_HEADERS.indexOf('Final Payment Amount')]  || 0);
    existingFPM  = String(ex[JOB_HEADERS.indexOf('Final Payment Method')]  || '');
    existingFPD  = String(ex[JOB_HEADERS.indexOf('Final Payment Date')]    || '');
    existingFPR  = String(ex[JOB_HEADERS.indexOf('Final Payment Ref')]     || '');
    existingBD   = Number(ex[JOB_HEADERS.indexOf('Balance Due')]           || 0);
    existingLR   = String(ex[JOB_HEADERS.indexOf('Lost Reason')]           || '');
    existingNotes= String(ex[JOB_HEADERS.indexOf('Notes')]                 || '');
    existingCEI  = String(ex[JOB_HEADERS.indexOf('Calendar Event ID')]     || '');
    existingUID  = String(ex[JOB_HEADERS.indexOf('Contact UID')]           || '').trim();   // v2.19.10
    existingName = String(ex[JOB_HEADERS.indexOf('Customer Name')]         || '').trim();   // v2.19.10
  }

  const inCP  = Number(data.contractPrice)     || 0;
  const inCPT = data.contractPriceType         || '';
  const inDP  = Number(data.depositPaid)       || 0;
  const inFPA = Number(data.finalPaymentAmount)|| 0;
  const inBD  = Number(data.balanceDue)        || 0;
  // v2.10.0 — cancelling a job's schedule needs to actually blank out the
  // schedule fields (including the calendar event ID), which the normal
  // `data.x || existingX` pattern can't do since an intentional '' always
  // falls back to the existing value. This flag is the explicit opt-in.
  const clearSchedule = !!data.clearSchedule;

  const row = [
    data.jobId||'', data.estimateId||'', data.contactUID||existingUID, data.custName||existingName,
    finalStatus, new Date().toISOString(), 'FencePro',
    inCP  > 0 ? inCP  : existingCP,
    inCPT      ? inCPT : existingCPT,
    data.depositType||'50pct',
    data.depositType === 'none' ? 0 : (Number(data.depositRequired)||0),
    inDP  > 0 ? inDP  : existingDP,
    data.depositMethod || existingDM,
    data.depositDate   || existingDD,
    data.depositRef    || existingDR,
    clearSchedule ? '' : (data.scheduledDate  || existingSD),
    clearSchedule ? '' : (data.scheduledCrew  || existingSC),
    clearSchedule ? '' : (data.scheduledTime  || existingST),
    clearSchedule ? '' : (data.scheduleNotes  || existingSN),
    inFPA > 0 ? inFPA : existingFPA,
    data.finalPaymentMethod || existingFPM,
    data.finalPaymentDate   || existingFPD,
    data.finalPaymentRef    || existingFPR,
    inBD  > 0 ? inBD  : existingBD,
    data.lostReason || existingLR,
    data.notes      || existingNotes,
    clearSchedule ? '' : (data.calendarEventId !== undefined ? data.calendarEventId : existingCEI)
  ];

  if (existingJobRow > 0) {
    sheet.getRange(existingJobRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  // ── Mirror payments to Payments sheet ──────────────────────
  // Write a Payments row only when the incoming value differs from what's
  // already on the Jobs row AND no matching active payment already exists
  // in the Payments sheet (dedup guard for retried requests).

  if (inDP > 0 && inDP !== existingDP) {
    // Check Payments sheet for an existing non-voided deposit of the same amount
    const _existingPmts = getPayments({ jobId: data.jobId });
    const _depAlready   = (_existingPmts.paymentMap[data.jobId]?.payments || [])
      .some(p => p.type === 'deposit' && p.amount === inDP);
    if (!_depAlready) {
      writePayment(
        data.jobId,
        'deposit',
        inDP,
        data.depositMethod || '',
        data.depositDate   || new Date().toISOString().split('T')[0],
        data.depositRef    || '',
        ''   // no invoice yet for a deposit
      );
    }
  }

  if (inFPA > 0 && inFPA !== existingFPA) {
    // Check for an existing non-voided final/partial payment of the same amount
    const _existingPmts2 = getPayments({ jobId: data.jobId });
    const _finAlready    = (_existingPmts2.paymentMap[data.jobId]?.payments || [])
      .some(p => (p.type === 'final' || p.type === 'partial') && p.amount === inFPA);
    if (!_finAlready) {
      writePayment(
        data.jobId,
        'final',
        inFPA,
        data.finalPaymentMethod || existingFPM,
        data.finalPaymentDate   || existingFPD || new Date().toISOString().split('T')[0],
        data.finalPaymentRef    || existingFPR,
        data.invoiceId || ''
      );
    }
  }

  return { success: true };
}

// ══════════════════════════════════════════════════════════════
// GET SINGLE JOB
// ══════════════════════════════════════════════════════════════
function getJob(data) {
  const estSheet = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const jobSheet = getOrCreateSheet(JOBS_TAB, JOB_HEADERS);
  const estData  = estSheet.getDataRange().getValues();
  const jobData  = jobSheet.getDataRange().getValues();
  if (estData.length <= 1) return { error: 'Not found' };

  const estHeaders = estData[0].map(h => String(h).trim());
  const jobHeaders = jobData.length > 1 ? jobData[0].map(h => String(h).trim()) : [];

  const lookupId = String(data.jobId||data.estimateId||'').trim();
  const estRow = estData.slice(1).find(row =>
    String(row[0]).trim() === lookupId || String(row[1]).trim() === lookupId
  );
  if (!estRow) return { error: 'Estimate not found' };

  const est = {};
  estHeaders.forEach((h, i) => { est[h] = sheetVal(estRow[i]); });

  // Always use numeric prefix — Jobs sheet col A is always the numeric prefix
  const jobId = String(est['Job ID']||est['Estimate ID']||'').split('-')[0].trim();
  let job = {};
  if (jobData.length > 1 && jobId) {
    const jobRows = jobData.slice(1).filter(r => String(r[0]).trim() === jobId);
    if (jobRows.length > 0) {
      jobHeaders.forEach((h, i) => { job[h] = sheetVal(jobRows[jobRows.length-1][i]); });
    }
  }

  const merged = { ...job, ...est };
  merged['Status'] = String(est['Status']||'').trim() || String(job['Status']||'').trim() || 'Draft';

  // Same Jobs-win logic as getJobs
  // Contract Price: prefer non-zero from either source
  const cpJobS  = Number(job['Contract Price'])  || 0;
  const cpEstS  = Number(est['Contract Price'])  || 0;
  const cptJobS = String(job['Contract Price Type'] || '');
  const cptEstS = String(est['Contract Price Type'] || '');
  merged['Contract Price']      = cpJobS  > 0 ? cpJobS  : (cpEstS  > 0 ? cpEstS  : 0);
  merged['Contract Price Type'] = cptJobS      ? cptJobS : (cptEstS      ? cptEstS : '');

  const JOBS_WIN_FIELDS = [
    'Deposit Type','Deposit Required',
    'Deposit Paid','Deposit Method','Deposit Date','Deposit Ref',
    'Balance Due','Scheduled Date','Scheduled Crew','Scheduled Time','Schedule Notes',
    'Final Payment Amount','Final Payment Method','Final Payment Date','Final Payment Ref',
    'Lost Reason','Calendar Event ID'
  ];
  JOBS_WIN_FIELDS.forEach(field => {
    const jobVal = job[field];
    const jobHasData = jobVal !== undefined && jobVal !== null && jobVal !== '' && jobVal !== 0;
    if (jobHasData) merged[field] = jobVal;
    else if (est[field] !== undefined) merged[field] = est[field];
  });

  // Attach payments from Payments sheet
  const pmtResult  = getPayments({ jobId: jobId });
  const pmtData    = pmtResult.paymentMap[jobId] || null;
  if (pmtData) {
    merged.payments        = pmtData.payments;
    merged.totalPaid       = pmtData.totalPaid;
    merged.depositPaid     = pmtData.depositPaid;
    merged.finalPaid       = pmtData.finalPaid;
    const cp               = merged['Contract Price'] || 0;
    merged.balanceDue_calc = cp > 0 ? Math.max(0, cp - pmtData.totalPaid) : 0;
  } else {
    merged.payments        = [];
    merged.totalPaid       = Number(merged['Deposit Paid'] || 0) + Number(merged['Final Payment Amount'] || 0);
    merged.depositPaid     = Number(merged['Deposit Paid'] || 0);
    merged.finalPaid       = Number(merged['Final Payment Amount'] || 0);
    merged.balanceDue_calc = merged['Balance Due'] || 0;
  }

  return { job: merged };
}

// ══════════════════════════════════════════════════════════════
// DELETE JOB — removes estimate from Estimates sheet.
// If it's the last estimate for the job, also removes from:
//   Jobs, Invoices, Payments, Proposal_Tokens, Job_Costs
function deleteJob(data) {
  const estSheet  = getOrCreateSheet(ESTIMATES_TAB,  ESTIMATE_HEADERS);
  const jobSheet  = getOrCreateSheet(JOBS_TAB,        JOB_HEADERS);
  const invSheet  = getOrCreateSheet(INVOICES_TAB,    INVOICE_HEADERS);
  const paySheet  = getOrCreateSheet(PAYMENTS_TAB,    PAYMENT_HEADERS);
  const tokSheet  = getOrCreateSheet(TOKENS_TAB,      TOKEN_HEADERS);
  const costSheet = getOrCreateSheet(JOB_COSTS_TAB,   JOB_COST_HEADERS);

  const estimateId = String(data.estimateId || data.jobId || '').trim();
  const jobId      = String(data.jobId || '').trim() || estimateId.split('-')[0];
  if (!estimateId) return { error: 'No estimateId provided' };

  // 1. Delete from Estimates
  let estData = estSheet.getDataRange().getValues();
  for (let i = estData.length - 1; i >= 1; i--) {
    if (String(estData[i][0]).trim() === estimateId) estSheet.deleteRow(i + 1);
  }

  // 2. Check if any estimates remain for this job
  estData = estSheet.getDataRange().getValues();
  const _estHdrs = estData[0].map(h => String(h).trim());
  const _estJobIdCol = _estHdrs.indexOf('Job ID');
  const remaining = estData.slice(1).filter(r => {
    const estIdVal = String(r[0] || '').trim();
    const jobIdVal = _estJobIdCol >= 0 ? String(r[_estJobIdCol] || '').trim() : '';
    return estIdVal.split('-')[0] === jobId || jobIdVal === jobId;
  });

  if (remaining.length === 0) {
    // Last estimate — purge all related records

    // Jobs sheet
    let d = jobSheet.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][0]).trim() === jobId) jobSheet.deleteRow(i + 1);
    }

    // Invoices sheet — match by Job ID (col 3, index 2) or Estimate ID (col 4, index 3)
    d = invSheet.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      const invJobId  = String(d[i][2] || '').trim();
      const invEstId  = String(d[i][3] || '').trim();
      const invPrefix = String(d[i][3] || d[i][2] || '').trim().split('-')[0];
      if (invJobId === jobId || invEstId === estimateId || invPrefix === jobId) {
        invSheet.deleteRow(i + 1);
      }
    }

    // Payments sheet — match by Job ID (col 2, index 1)
    d = paySheet.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1] || '').trim() === jobId) paySheet.deleteRow(i + 1);
    }

    // Proposal_Tokens — match by Estimate ID (col 2, index 1) or Estimate IDs list (col 14, index 13)
    d = tokSheet.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      const tokEstId  = String(d[i][1]  || '').trim();
      const tokEstIds = String(d[i][13] || '').trim();
      if (tokEstId === estimateId || tokEstIds.split(',').map(s => s.trim()).some(s => s.split('-')[0] === jobId)) {
        tokSheet.deleteRow(i + 1);
      }
    }

    // Job_Costs — match by Job ID (col 2, index 1)
    d = costSheet.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1] || '').trim() === jobId) costSheet.deleteRow(i + 1);
    }
  } else {
    // Siblings remain — only delete costs and tokens tied to this specific estimate
    let d = tokSheet.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1] || '').trim() === estimateId) tokSheet.deleteRow(i + 1);
    }
    d = costSheet.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][2] || '').trim() === estimateId) costSheet.deleteRow(i + 1);
    }
  }

  return { success: true, deleted: estimateId, jobPurged: remaining.length === 0 };
}

// ══════════════════════════════════════════════════════════════
// GET INVOICES
// ══════════════════════════════════════════════════════════════
function getInvoices() {
  try {
    const sheet   = getOrCreateSheet(INVOICES_TAB, INVOICE_HEADERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { invoices: [] };
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const invoices = data.slice(1).filter(row => row[0]).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = row[i];
        if (v instanceof Date) v = isNaN(v.getTime()) ? '' : v.toISOString().split('T')[0];
        obj[h] = v;
      });
      return obj;
    }).reverse();
    return { invoices, count: invoices.length };
  } catch(e) {
    return { invoices: [], error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// UPDATE INVOICE
// ══════════════════════════════════════════════════════════════
// v3.20.5 — the frontend has been calling action 'updateInvoice' (from
// confirmFinalPayment() and the Edit Invoice screen) for a while, but this
// action never existed here — it fell into doPost()'s default case and
// returned { error: 'Unknown action' }, which both callers silently
// swallow. Net effect: no invoice's Payment Status/Amount Due has ever
// actually been written back to the Invoices sheet, no matter how many
// times a job got marked paid — the Invoices tab always re-fetches the
// sheet fresh, so it kept showing the invoice's original status forever.
function updateInvoice(data) {
  const invoiceId = String((data && (data['Invoice ID'] || data.invoiceId)) || '').trim();
  if (!invoiceId) return { error: 'Missing invoiceId' };

  const sheet = getOrCreateSheet(INVOICES_TAB, INVOICE_HEADERS);
  const row   = findRow(sheet, invoiceId, INVOICE_HEADERS.indexOf('Invoice ID'));
  if (row < 0) return { error: 'Invoice not found: ' + invoiceId };

  // Pick the Title-Case key if present, else the camelCase one — but treat
  // 0 as a real, meaningful value (Amount Due: 0 means paid in full) rather
  // than falling through to the other key just because 0 is falsy.
  const pick = (a, b) => (a !== undefined && a !== null && a !== '') ? a : b;
  const FIELD_MAP = {
    'Payment Status': pick(data['Payment Status'], data.status),
    'Amount Due':     pick(data['Amount Due'],     data.amountDue),
    'Payment Method': pick(data['Payment Method'], data.paymentMethod),
    'Payment Date':   pick(data['Payment Date'],   data.paymentDate),
  };
  Object.keys(FIELD_MAP).forEach(col => {
    const val = FIELD_MAP[col];
    if (val === undefined || val === null || val === '') return;
    const colIdx = INVOICE_HEADERS.indexOf(col);
    if (colIdx >= 0) sheet.getRange(row, colIdx + 1).setValue(val);
  });

  return { success: true, invoiceId };
}

// ══════════════════════════════════════════════════════════════
// DELETE INVOICE
// ══════════════════════════════════════════════════════════════
// v2.8.2 — same "the frontend has been calling this for a while and it
// never existed" problem as updateInvoice, found in the same audit. The
// Invoices tab's delete button relied entirely on a local tombstone list
// (localStorage) to keep a deleted invoice from reappearing on THIS
// browser after a re-fetch from the sheet — but the row itself was never
// actually removed from the Invoices sheet, so it would still show up
// on any other device/browser, forever.
function deleteInvoice(data) {
  const invoiceId = String((data && (data.invoiceId || data['Invoice ID'])) || '').trim();
  if (!invoiceId) return { error: 'Missing invoiceId' };
  const sheet = getOrCreateSheet(INVOICES_TAB, INVOICE_HEADERS);
  const row   = findRow(sheet, invoiceId, INVOICE_HEADERS.indexOf('Invoice ID'));
  if (row < 0) return { error: 'Invoice not found: ' + invoiceId };
  sheet.deleteRow(row);
  return { success: true, deleted: invoiceId };
}

// ══════════════════════════════════════════════════════════════
// SAVE MANUAL INVOICE  (v2.13.0)
// ══════════════════════════════════════════════════════════════
// An invoice that belongs to no job. Line items, payments and sales tax are
// stored on the invoice row, so nothing here reads or writes the Jobs or
// Payments sheets. Create: no invoiceId. Edit: pass the invoiceId.
const MANUAL_INVOICE_FIRST = 5000;

// All money in whole cents so floating point can never drift a total.
function _miMoney(data) {
  const items = (Array.isArray(data.items) ? data.items : []).map(it => ({
    item:    String(it.item || '').trim(),
    desc:    String(it.desc || '').trim(),
    amount:  Math.round((Number(it.amount) || 0) * 100) / 100,
    taxable: it.taxable !== false
  })).filter(it => it.item || it.desc || it.amount);
  let subC = 0, taxableC = 0;
  items.forEach(it => { const c = Math.round(it.amount * 100); subC += c; if (it.taxable) taxableC += c; });
  let rate = Number(data.taxRate) || 0;
  if (!(rate >= 0 && rate <= 0.15)) rate = 0;
  const county = (['Suffolk', 'Nassau', 'Other'].indexOf(data.taxCounty) >= 0 && rate > 0) ? data.taxCounty : 'None';
  if (county === 'None') rate = 0;
  const taxC   = Math.max(0, Math.round(taxableC * rate + 1e-9));
  const totalC = subC + taxC;
  const pays = (Array.isArray(data.payments) ? data.payments : []).map(p => ({
    date:   String(p.date || '').slice(0, 10),
    label:  String(p.label || '').trim() || 'Payment',
    method: String(p.method || '').trim(),
    amount: Math.round((Number(p.amount) || 0) * 100) / 100
  })).filter(p => p.amount > 0);
  let paidC = 0;
  pays.forEach(p => { paidC += Math.round(p.amount * 100); });
  const balC   = Math.max(0, totalC - paidC);
  const status = (totalC > 0 && balC === 0) ? 'Paid in Full' : (paidC > 0 ? 'Partial' : 'Unpaid');
  return { items: items, pays: pays, rate: rate, county: county,
           subtotal: subC / 100, tax: taxC / 100, total: totalC / 100,
           paid: paidC / 100, balance: balC / 100, status: status };
}

// YYYY-MM-DD plus n days, with no time-zone involvement.
function _miAddDays(ymd, n) {
  const m = String(ymd || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return '';
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3] + n));
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
}

// Next manual number: never reuses one, even after the latest invoice is deleted.
function _miNextNumber(sheet, hdrRow) {
  let maxExisting = MANUAL_INVOICE_FIRST - 1;
  const typeC = hdrRow.indexOf('Invoice Type'), numC = hdrRow.indexOf('Invoice Number');
  if (typeC >= 0 && numC >= 0 && sheet.getLastRow() > 1) {
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, hdrRow.length).getValues();
    vals.forEach(r => {
      if (String(r[typeC]) === 'Manual') { const n = parseInt(r[numC], 10); if (n > maxExisting) maxExisting = n; }
    });
  }
  let seq = MANUAL_INVOICE_FIRST - 1;
  const settingsSheet = getOrCreateSheet(SETTINGS_TAB, ['Key', 'Value']);
  const sData = settingsSheet.getDataRange().getValues();
  let seqRow = -1;
  for (let i = 1; i < sData.length; i++) {
    if (String(sData[i][0]).trim() === 'manualInvoiceSeq') { seqRow = i + 1; seq = Number(sData[i][1]) || seq; break; }
  }
  const next = Math.max(maxExisting, seq) + 1;
  if (seqRow > 0) settingsSheet.getRange(seqRow, 2).setValue(next);
  else            settingsSheet.appendRow(['manualInvoiceSeq', next]);
  return next;
}

function saveManualInvoice(data) {
  data = data || {};
  if (!String(data.custName || '').trim()) return { error: 'Customer name is required' };
  const m = _miMoney(data);
  if (!m.items.length)  return { error: 'Add at least one item' };
  if (!(m.total > 0))   return { error: 'The invoice total must be greater than $0' };
  const invoiceDate = String(data.invoiceDate || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(invoiceDate)) return { error: 'Invoice date is required' };
  const dueDate = /^\d{4}-\d{2}-\d{2}$/.test(String(data.dueDate || '').slice(0, 10))
    ? String(data.dueDate).slice(0, 10) : _miAddDays(invoiceDate, 5);

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const sheet  = getOrCreateSheet(INVOICES_TAB, INVOICE_HEADERS);
    const hdrRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const colOf  = n => hdrRow.indexOf(n);

    let rowIdx = -1, invoiceId, invNum, row;
    if (data.invoiceId) {
      rowIdx = findRow(sheet, data.invoiceId, colOf('Invoice ID'));
      if (rowIdx < 0) return { error: 'Invoice not found: ' + data.invoiceId };
      row = sheet.getRange(rowIdx, 1, 1, hdrRow.length).getValues()[0];
      if (String(row[colOf('Invoice Type')]) !== 'Manual') return { error: 'That invoice is not a manual invoice' };
      invoiceId = String(data.invoiceId).trim();
      invNum    = String(row[colOf('Invoice Number')]);
    } else {
      invNum    = String(_miNextNumber(sheet, hdrRow));
      invoiceId = 'INV-' + invNum;
      row       = new Array(hdrRow.length).fill('');
    }
    const set = (name, v) => { const c = colOf(name); if (c >= 0) row[c] = v; };

    const last = m.pays.length ? m.pays[m.pays.length - 1] : null;
    set('Invoice ID', invoiceId);       set('Invoice Number', invNum);
    set('Job ID', 'MANUAL');            set('Job Prefix', '');
    set('Contact UID', String(data.contactUID || ''));
    set('Customer Name', String(data.custName).trim());
    set('Email', String(data.custEmail || '').trim());
    set('Customer Phone', String(data.custPhone || '').trim());
    set('Job Address', String(data.jobAddr || '').trim());
    set('Job Address 2', String(data.jobAddr2 || '').trim());
    set('Job City', String(data.jobCity || '').trim());
    set('Job State', String(data.jobState || 'NY').trim());
    set('Job Zip', String(data.jobZip || '').trim());
    set('Invoice Date', invoiceDate);   set('Due Date', dueDate);
    set('Fence Type', m.items[0].item || 'Manual Invoice');
    set('Final Price CC', m.total);     set('Cash Price', m.total);
    set('Customer Tax', m.tax);         set('Amount Due', m.balance);
    set('Payment Status', m.status);
    set('Payment Date', last ? last.date : '');
    set('Payment Method', last ? last.method : '');
    set('Payment Instructions', String(data.payNote || ''));
    set('Customer Note', String(data.customerNote || ''));
    set('Company Name', String(data.companyName || ''));
    set('Company Address', String(data.companyAddr || ''));
    set('Company Phone', String(data.companyPhone || ''));
    set('Company', String(data.company || ''));
    set('Invoice Type', 'Manual');
    set('Line Items', JSON.stringify(m.items));
    set('Payments', JSON.stringify(m.pays));
    set('Subtotal', m.subtotal);        set('Tax County', m.county);
    set('Tax Rate', m.rate);            set('Invoice Total', m.total);

    // Dates are written as plain text so Sheets never turns them into Date
    // cells (which shift a day when read back through toISOString).
    const target = rowIdx > 0 ? rowIdx : sheet.getLastRow() + 1;
    ['Invoice Date', 'Due Date', 'Payment Date'].forEach(n => {
      const c = colOf(n); if (c >= 0) sheet.getRange(target, c + 1).setNumberFormat('@');
    });
    sheet.getRange(target, 1, 1, hdrRow.length).setValues([row]);

    const invoice = {};
    hdrRow.forEach((h, i) => { invoice[h] = row[i]; });
    return { success: true, invoiceId: invoiceId, invoiceNumber: invNum, status: m.status,
             total: m.total, amountDue: m.balance, invoice: invoice };
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════
// SAVE INVOICE
// ══════════════════════════════════════════════════════════════
function saveInvoice(data) {
  const sheet = getOrCreateSheet(INVOICES_TAB, INVOICE_HEADERS);
  const year  = new Date().getFullYear();

  // ── Collision-safe invoice number ─────────────────────────────
  // Store a running sequence counter in Settings tab rather than
  // deriving from lastRow — avoids duplicate numbers on concurrent saves.
  let invSeq = 1;
  try {
    const settingsSheet = getOrCreateSheet(SETTINGS_TAB, ['Key','Value']);
    const sData = settingsSheet.getDataRange().getValues();
    const seqKey = 'invoiceSeq_' + year;
    let seqRow = -1;
    for (let i = 1; i < sData.length; i++) {
      if (String(sData[i][0]).trim() === seqKey) { seqRow = i + 1; invSeq = Number(sData[i][1]) + 1; break; }
    }
    if (seqRow > 0) settingsSheet.getRange(seqRow, 2).setValue(invSeq);
    else            settingsSheet.appendRow([seqKey, invSeq]);
  } catch(e) {
    // Fall back to lastRow count if settings write fails
    invSeq = Math.max(1, sheet.getLastRow());
    Logger.log('saveInvoice: seq counter fallback — ' + e.toString());
  }

  const invNum    = String(invSeq).padStart(4,'0');
  const invoiceId = data.invoiceId || ('INV-'+year+'-'+invNum);
  const existingRow = data.invoiceId ? findRow(sheet, data.invoiceId, 0) : -1;

  const row = [
    invoiceId, invNum,
    data.jobId||'', data.jobPrefix||'', data.contactUID||'',
    data.custName||'', data.custEmail||'', data.jobAddr||'',
    data.jobCity||'', data.jobState||'', data.jobZip||'',
    data.invoiceDate||'', data.dueDate||'', data.fenceType||'',
    Number(data.finalPrice)||0,
    Number(data.cashPrice)||0,
    Number(data.custTax)||0,
    // Amount Due: use provided value, or calculate from contract minus deposit if 0
    (() => {
      const amt = Number(data.amountDue)||0;
      if (amt > 0) return amt;
      const cp = Number(data.finalPrice)||0;
      const dp = Number(data.depositPaid)||0;
      return cp > 0 ? Math.max(0, cp - dp) : 0;
    })(),
    data.status||'Unpaid',
    data.paymentDate||'', data.paymentMethod||'', data.notes||'',
    data.customerNote||'', '',
    // v2.8.0 — snapshot the selected company's real name/address/phone
    // at invoice-creation time (front end resolves these from the
    // Companies list; previously computed but had no column to land in).
    data.companyName||'', data.companyAddr||'', data.companyPhone||''
  ];

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return { success: true, invoiceId, invoiceNumber: invNum };
}

// ══════════════════════════════════════════════════════════════
// UPLOAD PHOTO
// ══════════════════════════════════════════════════════════════
function uploadPhoto(data) {
  try {
    const folder     = getJobFolder(data.jobId, data.custName) || getRootFolder();
    const photosSub  = getOrCreateSubfolder(folder, 'Estimate Photos');
    photosSub.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.base64Data),
      data.mimeType||'image/jpeg',
      data.fileName||'photo.jpg'
    );
    const file   = photosSub.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId = file.getId();
    const directUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
    return { success: true, url: directUrl, fileId, name: data.fileName };
  } catch(e) {
    return { error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// UPLOAD DOC
// ══════════════════════════════════════════════════════════════
function uploadDoc(data) {
  try {
    const jobFolder = getJobFolder(data.jobId, data.custName) || getRootFolder();
    const letter    = (data.estimateLetter||'A').toUpperCase();
    const estFolder = getOrCreateSubfolder(jobFolder, 'EST-'+letter);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.base64Data),
      data.mimeType||'application/octet-stream',
      data.fileName||'document'
    );
    const file = estFolder.createFile(blob);
    return { success: true, url: file.getUrl(), name: data.fileName };
  } catch(e) {
    return { error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// SETTINGS
// ══════════════════════════════════════════════════════════════
function saveSettings(data) {
  const sheet = getOrCreateSheet(SETTINGS_TAB, ['Key','Value']);
  const existing = sheet.getDataRange().getValues();
  const map = {};
  existing.forEach((row, i) => { if (i > 0 && row[0]) map[row[0]] = i+1; });
  const now = new Date().toISOString();
  const entries = { ...data, lastSaved: now };
  Object.entries(entries).forEach(([k, v]) => {
    if (map[k]) sheet.getRange(map[k], 2).setValue(v);
    else { sheet.appendRow([k, v]); map[k] = sheet.getLastRow(); }
  });
  return { success: true, savedAt: now };
}

function getSettings() {
  try {
    const sheet  = getOrCreateSheet(SETTINGS_TAB, ['Key','Value']);
    const data   = sheet.getDataRange().getValues();
    const result = {};
    data.slice(1).forEach(row => { if (row[0]) result[row[0]] = row[1]; });
    return { settings: result };
  } catch(e) {
    return { settings: {} };
  }
}

// ══════════════════════════════════════════════════════════════
// CONTACTS
// ══════════════════════════════════════════════════════════════
function contactToRow(c) {
  return [
    c.uid||'', c.name||'', c.phone||'', c.altPhone||'', c.email||'',
    c.company||'', c.category||'', c.nickname||'',
    c.addr1||'', c.addr2||'', c.city||'', c.state||'', c.zip||'', c.county||'', c.community||'',
    c.billAddr1||'', c.billAddr2||'', c.billCity||'', c.billState||'', c.billZip||'',
    c.channel||'', c.campaign||'', c.referredBy||'', c.tags||'',
    c.contactType||'', c.projectType||'', c.projectedLF||'', c.fenceKind||'', c.fenceStyle||'',
    c.services||'', c.gateCode||'',
    c.customerSince||'', c.lastServiceDate||'', c.lifetimeValue||'', c.balanceDue||'',
    c.taxRate||'', c.cardOnFile||'', c.autoPay||'', c.doNotEmail||'', c.doNotText||'',
    c.flashMessage||'', c.additionalContacts||'',
    c.createdAt||new Date().toISOString(), c.source||'FencePro', c.inactive?'Yes':'No',
    c.estApptDateTime||'', c.estApptNotes||'', c.estApptEventId||'',
    c.properties||''
  ];
}

function rowToContact(row) {
  return {
    uid:row[0]||'', name:row[1]||'', phone:row[2]||'', altPhone:row[3]||'', email:row[4]||'',
    company:row[5]||'', category:row[6]||'', nickname:row[7]||'',
    addr1:row[8]||'', addr2:row[9]||'', city:row[10]||'', state:row[11]||'', zip:row[12]||'', county:row[13]||'', community:row[14]||'',
    billAddr1:row[15]||'', billAddr2:row[16]||'', billCity:row[17]||'', billState:row[18]||'', billZip:row[19]||'',
    channel:row[20]||'', campaign:row[21]||'', referredBy:row[22]||'', tags:row[23]||'',
    contactType:row[24]||'', projectType:row[25]||'', projectedLF:row[26]||'', fenceKind:row[27]||'', fenceStyle:row[28]||'',
    services:row[29]||'', gateCode:row[30]||'',
    customerSince:row[31]||'', lastServiceDate:row[32]||'', lifetimeValue:row[33]||'', balanceDue:row[34]||'',
    taxRate:row[35]||'', cardOnFile:row[36]||'', autoPay:row[37]||'', doNotEmail:row[38]||'', doNotText:row[39]||'',
    flashMessage:row[40]||'', additionalContacts:row[41]||'',
    createdAt:row[42]||'', source:row[43]||'FencePro', inactive:row[44]==='Yes',
    estApptDateTime:row[45]||'', estApptNotes:row[46]||'', estApptEventId:row[47]||'',
    properties:row[48]||''
  };
}

function saveContact(data) {
  const sheet  = getOrCreateSheet(CONTACTS_TAB, CONTACT_HEADERS);

  // New/unknown zip entered on Contact form — persist it for future lookups.
  // Dedup-safe; no-ops if the zip is already known.
  saveZipEntry(data.zip, data.city, data.county);

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.uid)) {
      sheet.getRange(i+1,1,1,CONTACT_HEADERS.length).setValues([contactToRow(data)]);
      return { success:true, action:'updated' };
    }
  }
  sheet.appendRow(contactToRow(data));
  return { success:true, action:'created', uid:data.uid };
}

function getSheetContacts() {
  try {
    const sheet   = getOrCreateSheet(CONTACTS_TAB, CONTACT_HEADERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { contacts:[] };
    const values  = sheet.getRange(2,1,lastRow-1,CONTACT_HEADERS.length).getValues();
    return { contacts: values.filter(r=>r[0]||r[1]).map(r=>rowToContact(r)), total: lastRow-1 };
  } catch(e) {
    return { contacts:[], error:e.message };
  }
}

function bulkSaveContacts(data) {
  const contacts = data.contacts||[];
  if (!contacts.length) return { success:true, saved:0 };
  const sheet   = getOrCreateSheet(CONTACTS_TAB, CONTACT_HEADERS);
  const lastRow = sheet.getLastRow();
  const uidMap  = {};
  if (lastRow > 1) {
    sheet.getRange(2,1,lastRow-1,1).getValues().forEach((r,i) => { if(r[0]) uidMap[String(r[0])]=i+2; });
  }
  const toAppend = [];
  contacts.forEach(c => {
    const row = contactToRow(c);
    if (c.uid && uidMap[c.uid]) sheet.getRange(uidMap[c.uid],1,1,CONTACT_HEADERS.length).setValues([row]);
    else toAppend.push(row);
  });
  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow()+1,1,toAppend.length,CONTACT_HEADERS.length).setValues(toAppend);
  }
  return { success:true, saved:contacts.length };
}

function deleteContact(data) {
  const sheet  = getOrCreateSheet(CONTACTS_TAB, CONTACT_HEADERS);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length-1; i >= 1; i--) {
    if (String(values[i][0]) === String(data.uid)) {
      sheet.deleteRow(i+1);
      return { success:true };
    }
  }
  return { error:'Contact not found' };
}

function getContacts() { return getSheetContacts(); }

// ══════════════════════════════════════════════════════════════
// CONTACT LINK AUDIT  (v2.19.3, record-level detail added in v2.19.4)
// READ ONLY. Changes no data and creates no sheets.
// Reports how well Estimates, Jobs and Invoices tie back to a Contact UID,
// and which contacts look like duplicates. Sheets are read by header name,
// never by column position. Row numbers are the real sheet row numbers.
// To run it: in the Apps Script editor pick runContactLinkAudit, press Run,
// then open Execution log. It can also be called through the
// getContactLinkAudit action.
// ══════════════════════════════════════════════════════════════
function getContactLinkAudit() {
  try {
    const ss  = getSpreadsheet();
    const CAP = 40;   // max sample rows listed per category

    const norm    = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
    const phone10 = s => { const d = String(s == null ? '' : s).replace(/\D/g, ''); return d.length >= 10 ? d.slice(-10) : ''; };

    // Reads a whole tab without creating or changing anything.
    function readTable(tab) {
      const sh = ss.getSheetByName(tab);
      if (!sh || sh.getLastRow() < 1) return null;
      const vals = sh.getDataRange().getValues();
      const idx  = {};
      vals[0].forEach((h, i) => { const k = String(h).trim(); if (k && !(k in idx)) idx[k] = i; });
      return { idx: idx, rows: vals.slice(1) };   // rows[i] is sheet row i + 2
    }
    const cell = (row, idx, hdr) => (hdr in idx) ? String(row[idx[hdr]] == null ? '' : row[idx[hdr]]).trim() : '';

    // ── Contacts ──
    const ct = readTable(CONTACTS_TAB);
    if (!ct) return { error: 'Contacts sheet not found or empty' };

    const byUid     = Object.create(null);   // uid  -> contact name
    const uidCount  = Object.create(null);   // uid  -> how many contact rows use it
    const rowsByUid = Object.create(null);   // uid  -> [row details]
    const byName    = Object.create(null);   // normalized name  -> [uid, ...]
    const byPhone   = Object.create(null);
    const byEmail   = Object.create(null);
    let total = 0, blankUid = 0, inactive = 0;

    ct.rows.forEach((r, i) => {
      const uid  = cell(r, ct.idx, 'UID');
      const name = cell(r, ct.idx, 'Name');
      if (!uid && !name) return;
      total++;
      if (cell(r, ct.idx, 'Inactive') === 'Yes') inactive++;
      if (!uid) blankUid++;
      else {
        uidCount[uid] = (uidCount[uid] || 0) + 1;
        byUid[uid] = name;
        (rowsByUid[uid] = rowsByUid[uid] || []).push({
          row: i + 2, name: name,
          phone: cell(r, ct.idx, 'Phone'), email: cell(r, ct.idx, 'Email'),
          addr1: cell(r, ct.idx, 'Address 1'), createdAt: cell(r, ct.idx, 'Created At'),
          source: cell(r, ct.idx, 'Source')
        });
      }

      const add = (map, key) => { if (!key) return; (map[key] = map[key] || []).push(uid); };
      add(byName,  norm(name));
      add(byPhone, phone10(cell(r, ct.idx, 'Phone')));
      add(byEmail, norm(cell(r, ct.idx, 'Email')));
    });

    function dupGroups(map) {
      const keys = Object.keys(map).filter(k => map[k].length > 1);
      return { groups: keys.length, samples: keys.slice(0, CAP).map(k => ({ key: k, count: map[k].length, uids: map[k] })) };
    }
    const dupUids = Object.keys(uidCount).filter(u => uidCount[u] > 1)
      .map(u => ({ uid: u, count: uidCount[u], rows: rowsByUid[u] }));

    // ── Tables that should carry a Contact UID ──
    const specs = [
      { label: 'Estimates', tab: ESTIMATES_TAB, id: 'Estimate ID' },
      { label: 'Jobs',      tab: JOBS_TAB,      id: 'Job ID' },
      { label: 'Invoices',  tab: INVOICES_TAB,  id: 'Invoice ID' }
    ];
    const tables = {};

    specs.forEach(sp => {
      const t = readTable(sp.tab);
      if (!t) { tables[sp.label] = { missing: true }; return; }
      if (!('Contact UID' in t.idx)) { tables[sp.label] = { error: 'No Contact UID column found' }; return; }

      const out = {
        rows: 0, linked: 0, linkedNameDiffers: 0,
        blankUid: 0, blankOneNameMatch: 0, blankNoNameMatch: 0, blankSeveralNameMatches: 0,
        orphanUid: 0, orphanOneNameMatch: 0,
        samples: { blankOneNameMatch: [], blankNoNameMatch: [], blankSeveralNameMatches: [], orphanUid: [], linkedNameDiffers: [] }
      };
      const push = (arr, o) => { if (arr.length < CAP) arr.push(o); };

      t.rows.forEach((r, i) => {
        const id   = cell(r, t.idx, sp.id);
        const name = cell(r, t.idx, 'Customer Name');
        if (!id && !name) return;
        out.rows++;

        const row     = i + 2;
        const uid     = cell(r, t.idx, 'Contact UID');
        const matches = (byName[norm(name)] || []).filter(Boolean);   // UIDs of contacts with this exact name

        if (!uid) {
          out.blankUid++;
          if (matches.length === 1) {
            out.blankOneNameMatch++;
            push(out.samples.blankOneNameMatch, { row: row, id: id, name: name, linkTo: matches[0], contactName: byUid[matches[0]] });
          } else if (matches.length === 0) {
            out.blankNoNameMatch++;
            push(out.samples.blankNoNameMatch, { row: row, id: id, name: name });
          } else {
            out.blankSeveralNameMatches++;
            push(out.samples.blankSeveralNameMatches, { row: row, id: id, name: name, candidates: matches });
          }
        } else if (!(uid in byUid)) {
          out.orphanUid++;
          if (matches.length === 1) out.orphanOneNameMatch++;
          push(out.samples.orphanUid, { row: row, id: id, name: name, uid: uid, nameMatchUids: matches });
        } else {
          out.linked++;
          if (norm(byUid[uid]) !== norm(name)) {
            out.linkedNameDiffers++;
            push(out.samples.linkedNameDiffers, { row: row, id: id, name: name, uid: uid, contactName: byUid[uid], nameMatchUids: matches });
          }
        }
      });
      tables[sp.label] = out;
    });

    return {
      success: true,
      generatedAt: new Date().toISOString(),
      contacts: {
        total: total, inactive: inactive, blankUid: blankUid,
        duplicateUids: dupUids,
        duplicateNames:  dupGroups(byName),
        duplicatePhones: dupGroups(byPhone),
        duplicateEmails: dupGroups(byEmail)
      },
      tables: tables
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// Editor-friendly runner. Prints a short summary, then record-level detail.
// Each section is its own log entry so no single entry is too large.
function runContactLinkAudit() {
  const r = getContactLinkAudit();
  if (r.error) { Logger.log('AUDIT ERROR: ' + r.error); return r; }

  const q = s => '"' + s + '"';
  const L = [];
  L.push('CONTACT LINK AUDIT  ' + r.generatedAt);
  L.push('');
  L.push('CONTACTS: ' + r.contacts.total + ' total, ' + r.contacts.inactive + ' inactive');
  L.push('  blank UID: ' + r.contacts.blankUid + '   UIDs used more than once: ' + r.contacts.duplicateUids.length);
  L.push('  duplicate groups  name: ' + r.contacts.duplicateNames.groups +
         '   phone: ' + r.contacts.duplicatePhones.groups +
         '   email: ' + r.contacts.duplicateEmails.groups);
  Object.keys(r.tables).forEach(label => {
    const t = r.tables[label];
    L.push('');
    if (t.missing) { L.push(label.toUpperCase() + ': sheet not found'); return; }
    if (t.error)   { L.push(label.toUpperCase() + ': ' + t.error); return; }
    L.push(label.toUpperCase() + ': ' + t.rows + ' rows');
    L.push('  linked to a real contact: ' + t.linked + '  (name differs from contact: ' + t.linkedNameDiffers + ')');
    L.push('  blank Contact UID: ' + t.blankUid +
           '  -> exactly one contact with that name: ' + t.blankOneNameMatch +
           ', no match: ' + t.blankNoNameMatch +
           ', several matches: ' + t.blankSeveralNameMatches);
    L.push('  UID not found in Contacts: ' + t.orphanUid + '  (of those, one contact with that name: ' + t.orphanOneNameMatch + ')');
  });
  Logger.log(L.join('\n'));

  // Contact UIDs that appear on more than one contact row
  if (r.contacts.duplicateUids.length) {
    const D = ['DUPLICATE CONTACT UIDs (same UID on more than one Contacts row):'];
    r.contacts.duplicateUids.forEach(d => {
      D.push('  UID ' + d.uid + ' is on ' + d.count + ' rows:');
      d.rows.forEach(x => D.push('    row ' + x.row + ' | ' + x.name + ' | ' + x.phone + ' | ' + x.email +
                                 ' | ' + x.addr1 + ' | created ' + x.createdAt + ' | source ' + x.source));
    });
    Logger.log(D.join('\n'));
  }

  // Record-level detail for each table
  Object.keys(r.tables).forEach(label => {
    const t = r.tables[label];
    if (!t || !t.samples) return;
    const S = t.samples;
    const D = [];
    const section = (title, total, arr, fmt) => {
      if (!total) return;
      D.push('  ' + title + ' (' + total + (arr.length < total ? ', showing first ' + arr.length : '') + '):');
      arr.forEach(x => D.push('    ' + fmt(x)));
    };
    section('BLANK UID, exactly one contact with that name (safe to link)', t.blankOneNameMatch, S.blankOneNameMatch,
      x => 'row ' + x.row + ' | ' + x.id + ' | ' + q(x.name) + ' -> would link to ' + x.linkTo + ' (' + x.contactName + ')');
    section('BLANK UID, no contact with that name', t.blankNoNameMatch, S.blankNoNameMatch,
      x => 'row ' + x.row + ' | ' + x.id + ' | ' + q(x.name));
    section('BLANK UID, several contacts with that name', t.blankSeveralNameMatches, S.blankSeveralNameMatches,
      x => 'row ' + x.row + ' | ' + x.id + ' | ' + q(x.name) + ' | candidates: ' + x.candidates.join(', '));
    section('UID NOT FOUND IN CONTACTS', t.orphanUid, S.orphanUid,
      x => 'row ' + x.row + ' | ' + x.id + ' | ' + q(x.name) + ' | missing UID ' + x.uid +
           ' | contacts with this name: ' + (x.nameMatchUids.length ? x.nameMatchUids.join(', ') : 'none'));
    section('LINKED BUT NAME DIFFERS (rename, or wrong link)', t.linkedNameDiffers, S.linkedNameDiffers,
      x => 'row ' + x.row + ' | ' + x.id + ' | record says ' + q(x.name) + ' | linked contact is ' + q(x.contactName) +
           ' (' + x.uid + ') | record name matches other contact: ' +
           (x.nameMatchUids.length ? x.nameMatchUids.join(', ') : 'none'));
    if (D.length) Logger.log(label.toUpperCase() + ' DETAIL\n' + D.join('\n'));
  });
  return r;
}

// ══════════════════════════════════════════════════════════════
// JOBS REPAIR PREVIEW  (v2.19.5)
// READ ONLY. Changes no data and creates no sheets.
// For every Jobs row the audit flagged, shows the full row, the matching
// Estimates rows, and what a repair WOULD change. Also scans every Jobs row
// for cells that look out of place (shifted columns), lists repeated Job IDs,
// and compares the cells of contact rows that share a UID.
// To run it: in the Apps Script editor pick previewJobsRepair, press Run,
// then open Execution log. Sheets are read by header name, never position.
// ══════════════════════════════════════════════════════════════
function previewJobsRepair() {
  const ss  = getSpreadsheet();
  const CAP = 60;
  const KNOWN_STATUS = ['Draft','Committed','Sent','Accepted','Deposit Received','Scheduled',
                        'In Progress','Ready for Invoice','Invoiced','Paid in Full','Lost'];
  const STATUS_LC = {};
  KNOWN_STATUS.forEach(s => { STATUS_LC[s.toLowerCase()] = true; });
  const NUMERIC_COLS = ['Contract Price','Deposit Required','Deposit Paid','Final Payment Amount','Balance Due'];

  const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  const fmt  = v => (v instanceof Date && !isNaN(v.getTime())) ? v.toISOString() : String(v == null ? '' : v).trim();
  const uniq = a => a.filter((x, i) => a.indexOf(x) === i);
  const clip = (s, n) => { s = String(s); return s.length > n ? s.substring(0, n) + '...' : s; };

  function readTable(tab) {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 1) return null;
    const vals = sh.getDataRange().getValues();
    const idx  = {};
    vals[0].forEach((h, i) => { const k = String(h).trim(); if (k && !(k in idx)) idx[k] = i; });
    return { headers: vals[0].map(h => String(h).trim()), idx: idx, rows: vals.slice(1) };   // rows[i] is sheet row i + 2
  }
  const cell = (row, idx, hdr) => (hdr in idx) ? fmt(row[idx[hdr]]) : '';

  const ct = readTable(CONTACTS_TAB);
  const et = readTable(ESTIMATES_TAB);
  const jt = readTable(JOBS_TAB);
  if (!ct || !et || !jt) { Logger.log('PREVIEW ERROR: Contacts, Estimates or Jobs sheet missing or empty'); return; }

  // ── Contact lookups ──
  const byUid  = Object.create(null);   // uid -> name
  const byName = Object.create(null);   // normalized name -> [uid]
  const contactRowsByUid = Object.create(null);
  ct.rows.forEach((r, i) => {
    const uid = cell(r, ct.idx, 'UID'), name = cell(r, ct.idx, 'Name');
    if (!uid && !name) return;
    if (uid) {
      byUid[uid] = name;
      (contactRowsByUid[uid] = contactRowsByUid[uid] || []).push({ row: i + 2, vals: r });
    }
    if (norm(name)) (byName[norm(name)] = byName[norm(name)] || []).push(uid);
  });

  // ── Estimates, loaded once ──
  const ests = [];
  et.rows.forEach((er, i) => {
    const id = cell(er, et.idx, 'Estimate ID');
    if (!id) return;
    ests.push({
      row: i + 2, id: id,
      job: cell(er, et.idx, 'Job ID'), pref: cell(er, et.idx, 'Job Prefix'), no: cell(er, et.idx, 'Job No'),
      uid: cell(er, et.idx, 'Contact UID'), name: cell(er, et.idx, 'Customer Name'), status: cell(er, et.idx, 'Status')
    });
  });
  function estimatesFor(jobRow) {
    const jid = cell(jobRow, jt.idx, 'Job ID');
    const eid = cell(jobRow, jt.idx, 'Estimate ID');
    return ests.filter(e => {
      if (eid && e.id === eid) return true;
      return !!jid && (e.id === jid || e.job === jid || e.pref === jid || e.no === jid || e.id.indexOf(jid + '-') === 0);
    });
  }
  const estLine = e => e.id + ' | "' + e.name + '" | uid ' + (e.uid || '(blank)') + ' | ' + (e.status || 'no status') + ' (Estimates row ' + e.row + ')';

  // Non-empty cells of a row as one readable line
  function rowCells(row, tbl) {
    const parts = [];
    tbl.headers.forEach((h, c) => { const v = fmt(row[c]); if (h && v !== '') parts.push(h + '=' + clip(v, 60)); });
    return parts.join(' | ');
  }
  // Headers whose values are not identical across a set of row arrays
  function differingHeaders(rowArrays, tbl) {
    const out = [];
    tbl.headers.forEach((h, c) => {
      const seen = uniq(rowArrays.map(r => fmt(r[c])));
      if (seen.length > 1) out.push(h || ('col ' + (c + 1)));
    });
    return out;
  }

  // ── Scan every Jobs row ──
  const blankRows = [], oddRows = [], suspectRows = [];
  const rowsByJobId = Object.create(null);
  let jobRowCount = 0;

  jt.rows.forEach((r, i) => {
    const row  = i + 2;
    const jid  = cell(r, jt.idx, 'Job ID');
    const name = cell(r, jt.idx, 'Customer Name');
    if (!jid && !name) return;
    jobRowCount++;
    const uid = cell(r, jt.idx, 'Contact UID');
    if (jid) (rowsByJobId[jid] = rowsByJobId[jid] || []).push({ row: row, vals: r });

    if (!uid) blankRows.push({ row: row, vals: r, jid: jid, name: name });

    const reasons = [];
    if (uid && !(uid in byUid)) reasons.push('UID not found in Contacts');
    if (uid && (uid in byUid) && norm(byUid[uid]) !== norm(name)) reasons.push('name differs from linked contact');
    if (STATUS_LC[norm(name)]) reasons.push('status word in the Customer Name cell');
    if (reasons.length) oddRows.push({ row: row, vals: r, jid: jid, name: name, uid: uid, reasons: reasons });

    // cells that do not belong in their column
    const bad = [];
    const st = cell(r, jt.idx, 'Status');
    if (st && KNOWN_STATUS.indexOf(st) < 0) bad.push('Status="' + clip(st, 30) + '"');
    if ('Status Changed At' in jt.idx) {
      const raw = r[jt.idx['Status Changed At']];
      if (raw && !(raw instanceof Date) && isNaN(Date.parse(String(raw)))) bad.push('Status Changed At="' + clip(String(raw), 30) + '"');
    }
    NUMERIC_COLS.forEach(col => {
      const v = cell(r, jt.idx, col);
      if (v !== '' && isNaN(Number(v.replace(/[$,]/g, '')))) bad.push(col + '="' + clip(v, 30) + '"');
    });
    if (bad.length) suspectRows.push({ row: row, jid: jid, name: name, bad: bad });
  });

  // ── Verdicts for blank-UID jobs ──
  const blankOut = [];
  const verdictCount = { SAFE: 0, REVIEW: 0, CONFLICT: 0, UNRESOLVED: 0 };
  blankRows.forEach(b => {
    const es       = estimatesFor(b.vals);
    const estUids  = uniq(es.map(e => e.uid).filter(u => u && (u in byUid)));
    const nameUids = uniq((byName[norm(b.name)] || []).filter(Boolean));
    let kind, note;
    if (estUids.length === 1 && nameUids.length === 1 && estUids[0] === nameUids[0]) {
      kind = 'SAFE'; note = 'set Contact UID = ' + estUids[0] + ' (' + byUid[estUids[0]] + '); estimate and name agree';
    } else if (estUids.length === 1 && nameUids.length === 1) {
      kind = 'CONFLICT'; note = 'estimates say ' + estUids[0] + ' (' + byUid[estUids[0]] + ') but name match is ' + nameUids[0];
    } else if (estUids.length === 1 && nameUids.length === 0) {
      kind = 'REVIEW'; note = 'estimates say ' + estUids[0] + ' (' + byUid[estUids[0]] + '); no contact has this exact name';
    } else if (estUids.length === 0 && nameUids.length === 1) {
      kind = 'REVIEW'; note = 'name match only (' + nameUids[0] + '), no estimate confirms it' + (es.length ? '' : '; no estimate found for this job');
    } else if (estUids.length > 1) {
      kind = 'REVIEW'; note = 'estimates point to several contacts: ' + estUids.join(', ');
    } else {
      kind = 'UNRESOLVED'; note = 'no usable contact from estimates or name';
    }
    verdictCount[kind]++;
    blankOut.push('  [' + kind + '] row ' + b.row + ' | job ' + b.jid + ' | "' + b.name + '" | ' + note);
  });

  // ── Duplicate Job IDs ──
  const dupJobIds = Object.keys(rowsByJobId).filter(k => rowsByJobId[k].length > 1);
  // ── Duplicate contact UIDs ──
  const dupContactUids = Object.keys(contactRowsByUid).filter(k => contactRowsByUid[k].length > 1);

  // ══ OUTPUT: summary first, so it survives even if the log is cut ══
  const S = [];
  S.push('JOBS REPAIR PREVIEW (read only, nothing was changed)  ' + new Date().toISOString());
  S.push('');
  S.push('Jobs rows scanned: ' + jobRowCount);
  S.push('Blank Contact UID: ' + blankRows.length + '  -> SAFE ' + verdictCount.SAFE + ', REVIEW ' + verdictCount.REVIEW +
         ', CONFLICT ' + verdictCount.CONFLICT + ', UNRESOLVED ' + verdictCount.UNRESOLVED);
  S.push('Odd rows (bad UID or name cell): ' + oddRows.length);
  S.push('Rows with out-of-place cells anywhere (full scan): ' + suspectRows.length);
  S.push('Job IDs that appear on more than one Jobs row: ' + dupJobIds.length);
  S.push('Contact UIDs on more than one Contacts row: ' + dupContactUids.length);
  Logger.log(S.join('\n'));

  // A. blank-UID jobs
  if (blankOut.length) Logger.log('A. BLANK CONTACT UID ON JOBS\n' + blankOut.join('\n'));

  // B. odd rows, one log entry per row
  oddRows.forEach(o => {
    const es       = estimatesFor(o.vals);
    const estNames = uniq(es.map(e => e.name).filter(Boolean));
    const estUids  = uniq(es.map(e => e.uid).filter(u => u && (u in byUid)));
    const contactName = (o.uid in byUid) ? byUid[o.uid] : '';
    const isStatusWord = !!STATUS_LC[norm(o.name)];
    const L = [];
    L.push('B. ODD ROW: Jobs row ' + o.row + ' | Job ID ' + o.jid);
    L.push('  why flagged: ' + o.reasons.join('; '));
    L.push('  current cells: ' + rowCells(o.vals, jt));
    L.push('  linked contact: ' + (contactName ? '"' + contactName + '" (' + o.uid + ')' : 'none (UID "' + o.uid + '" not found)'));
    L.push('  estimates for this job: ' + (es.length ? '' : 'none found'));
    es.forEach(e => L.push('    ' + estLine(e)));

    let proposal;
    if (!isStatusWord && estNames.length >= 1 && estNames.every(n => norm(n) === norm(o.name))) {
      proposal = 'NO JOBS REPAIR: the job name matches its estimate, so this is a typo on the estimate. Correct the name in the app.' +
                 (contactName ? ' Contact on file: "' + contactName + '".' : '');
    } else if (contactName && estNames.length === 1 && norm(estNames[0]) === norm(contactName)) {
      proposal = 'SAFE: set Customer Name to "' + contactName + '" (linked contact and estimate agree). Other cells in this row still need a look.';
    } else if (contactName && estNames.length === 1) {
      proposal = 'REVIEW: estimate says "' + estNames[0] + '", contact says "' + contactName + '".';
    } else if (contactName && estNames.length === 0) {
      proposal = 'REVIEW: no estimate found; contact says "' + contactName + '".';
    } else if (!contactName && estUids.length === 1) {
      proposal = 'REVIEW: estimates point to Contact UID ' + estUids[0] + ' ("' + byUid[estUids[0]] + '"). This row looks column-shifted, so check the other cells first.';
    } else {
      proposal = 'REVIEW: cannot decide from the data. Needs a manual look.';
    }
    L.push('  proposal: ' + proposal);
    Logger.log(L.join('\n'));
  });

  // C. out-of-place cells anywhere in Jobs
  if (suspectRows.length) {
    const L = ['C. OUT-OF-PLACE CELLS (full scan of every Jobs row)' + (suspectRows.length > CAP ? ', showing first ' + CAP + ' of ' + suspectRows.length : '')];
    suspectRows.slice(0, CAP).forEach(s => L.push('  row ' + s.row + ' | job ' + s.jid + ' | "' + s.name + '" | ' + s.bad.join(', ')));
    Logger.log(L.join('\n'));
  } else {
    Logger.log('C. OUT-OF-PLACE CELLS: none found in any Jobs row (statuses, status dates and money columns all look right)');
  }

  // D. repeated Job IDs
  if (dupJobIds.length) {
    const L = ['D. JOB IDs ON MORE THAN ONE JOBS ROW' + (dupJobIds.length > CAP ? ' (showing first ' + CAP + ' of ' + dupJobIds.length + ')' : '')];
    dupJobIds.slice(0, CAP).forEach(k => {
      const g = rowsByJobId[k];
      const diff = differingHeaders(g.map(x => x.vals), jt);
      L.push('  Job ID ' + k + ' on rows ' + g.map(x => x.row).join(', ') + ': ' +
             (diff.length ? 'rows differ in ' + diff.join(', ') : 'rows are identical in every cell'));
    });
    Logger.log(L.join('\n'));
  } else {
    Logger.log('D. JOB IDs: every Job ID appears on exactly one Jobs row');
  }

  // E. duplicate contact UIDs, all cells compared
  if (dupContactUids.length) {
    const L = ['E. CONTACT UIDs ON MORE THAN ONE CONTACTS ROW'];
    dupContactUids.forEach(u => {
      const g = contactRowsByUid[u];
      const diff = differingHeaders(g.map(x => x.vals), ct);
      L.push('  UID ' + u + ' ("' + byUid[u] + '") on rows ' + g.map(x => x.row).join(', ') + ': ' +
             (diff.length ? 'rows differ in ' + diff.join(', ') : 'identical in all ' + ct.headers.length + ' cells'));
    });
    Logger.log(L.join('\n'));
  }

  return { jobsScanned: jobRowCount, blank: blankRows.length, odd: oddRows.length, suspect: suspectRows.length };
}

// ══════════════════════════════════════════════════════════════
// JOBS REPAIR  (v2.19.6)
// Fixes ONLY the cases previewJobsRepair() showed as safe:
//   1. Jobs rows with a blank Contact UID, when the estimate and the customer
//      name agree on exactly one contact. Writes the Contact UID cell only.
//   2. Jobs rows whose Customer Name cell holds a status word (e.g. "Committed")
//      and whose Status cell holds the same word, when the linked contact and
//      the estimate agree on the real name. Writes the Customer Name cell only.
//   3. A row whose cells are all one column left of where they belong
//      (Contact UID cell holds a name, Customer Name cell holds a status,
//      Status cell holds a timestamp). The contact comes from the estimate.
//      The row is rewritten once with the missing UID inserted.
// Nothing else is touched. Every row is re-checked at the moment of the write
// and skipped if it no longer matches.
//
//   repairJobsDryRun() : prints exactly what would change. Changes nothing.
//   repairJobsApply()  : makes the changes and prints every old and new value.
//                        Save that log. It is your undo record.
// Back up first: right-click the Jobs tab, Duplicate.
// ══════════════════════════════════════════════════════════════
function _jobsRepairPlan() {
  const ss = getSpreadsheet();
  const KNOWN_STATUS = ['Draft','Committed','Sent','Accepted','Deposit Received','Scheduled',
                        'In Progress','Ready for Invoice','Invoiced','Paid in Full','Lost'];
  const STATUS_LC = {};
  KNOWN_STATUS.forEach(s => { STATUS_LC[s.toLowerCase()] = true; });
  const NUMERIC_COLS = ['Contract Price','Deposit Required','Deposit Paid','Final Payment Amount','Balance Due'];

  const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  const fmt  = v => (v instanceof Date && !isNaN(v.getTime())) ? v.toISOString() : String(v == null ? '' : v).trim();
  const uniq = a => a.filter((x, i) => a.indexOf(x) === i);
  const isDateLike = v => (v instanceof Date && !isNaN(v.getTime())) ||
                          (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v.trim()));

  function readTable(tab) {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 1) return null;
    const vals = sh.getDataRange().getValues();
    const idx  = {};
    vals[0].forEach((h, i) => { const k = String(h).trim(); if (k && !(k in idx)) idx[k] = i; });
    return { headers: vals[0].map(h => String(h).trim()), idx: idx, rows: vals.slice(1) };   // rows[i] is sheet row i + 2
  }
  const cell = (row, idx, hdr) => (hdr in idx) ? fmt(row[idx[hdr]]) : '';

  const ct = readTable(CONTACTS_TAB), et = readTable(ESTIMATES_TAB), jt = readTable(JOBS_TAB);
  if (!ct || !et || !jt) return { error: 'Contacts, Estimates or Jobs sheet missing or empty' };
  const need = ['Job ID', 'Estimate ID', 'Contact UID', 'Customer Name', 'Status', 'Status Changed At'];
  for (let k = 0; k < need.length; k++) {
    if (!(need[k] in jt.idx)) return { error: 'Jobs sheet is missing the "' + need[k] + '" column' };
  }

  const byUid = Object.create(null), byName = Object.create(null);
  ct.rows.forEach(r => {
    const uid = cell(r, ct.idx, 'UID'), name = cell(r, ct.idx, 'Name');
    if (!uid && !name) return;
    if (uid) byUid[uid] = name;
    if (norm(name)) (byName[norm(name)] = byName[norm(name)] || []).push(uid);
  });

  const ests = [];
  et.rows.forEach((er, i) => {
    const id = cell(er, et.idx, 'Estimate ID');
    if (!id) return;
    ests.push({ row: i + 2, id: id, job: cell(er, et.idx, 'Job ID'), pref: cell(er, et.idx, 'Job Prefix'),
                no: cell(er, et.idx, 'Job No'), uid: cell(er, et.idx, 'Contact UID'), name: cell(er, et.idx, 'Customer Name') });
  });
  function estimatesFor(jobRow) {
    const jid = cell(jobRow, jt.idx, 'Job ID'), eid = cell(jobRow, jt.idx, 'Estimate ID');
    return ests.filter(e => {
      if (eid && e.id === eid) return true;
      return !!jid && (e.id === jid || e.job === jid || e.pref === jid || e.no === jid || e.id.indexOf(jid + '-') === 0);
    });
  }

  const W = JOB_HEADERS.length;
  const changes = [], notTouched = [];

  jt.rows.forEach((r, i) => {
    const row  = i + 2;
    const jid  = cell(r, jt.idx, 'Job ID');
    const name = cell(r, jt.idx, 'Customer Name');
    if (!jid && !name) return;
    const uid    = cell(r, jt.idx, 'Contact UID');
    const status = cell(r, jt.idx, 'Status');
    const label  = 'job ' + jid + ' "' + name + '"';
    const es     = estimatesFor(r);

    // 1. blank Contact UID
    if (!uid) {
      const estUids  = uniq(es.map(e => e.uid).filter(u => u && (u in byUid)));
      const nameUids = uniq((byName[norm(name)] || []).filter(Boolean));
      if (estUids.length === 1 && nameUids.length === 1 && estUids[0] === nameUids[0]) {
        changes.push({ kind: 'uid', row: row, col: jt.idx['Contact UID'] + 1, header: 'Contact UID', oldVal: '', newVal: estUids[0], label: label });
      } else {
        notTouched.push('row ' + row + ' | ' + label + ' | blank Contact UID, estimate and name do not agree on one contact');
      }
      return;
    }

    // 3. whole row shifted one column left
    const looksShifted = !(uid in byUid) && !!STATUS_LC[norm(name)] &&
                         isDateLike(r[jt.idx['Status']]) &&
                         !!fmt(r[jt.idx['Status Changed At']]) && !isDateLike(r[jt.idx['Status Changed At']]);
    if (looksShifted) {
      const sameOrder = JOB_HEADERS.every((h, k) => jt.headers[k] === h);
      const exact     = ests.filter(e => e.id === cell(r, jt.idx, 'Estimate ID'));
      const estUids   = uniq(exact.map(e => e.uid).filter(u => u && (u in byUid)));
      const old       = r.slice(0, W);
      if (!sameOrder) {
        notTouched.push('row ' + row + ' | ' + label + ' | looks shifted, but the Jobs column order differs from the code, so it was not touched');
      } else if (estUids.length !== 1) {
        notTouched.push('row ' + row + ' | ' + label + ' | looks shifted, but the estimate does not give exactly one contact');
      } else if (old.length < W || fmt(old[W - 1]) !== '') {
        notTouched.push('row ' + row + ' | ' + label + ' | looks shifted, but the last cell is not blank, so shifting would lose data');
      } else {
        const nw = old.slice(0, 2).concat([estUids[0]], old.slice(2, W - 1));
        const okStatus = KNOWN_STATUS.indexOf(fmt(nw[jt.idx['Status']])) >= 0;
        const okDate   = isDateLike(nw[jt.idx['Status Changed At']]);
        const okNums   = NUMERIC_COLS.every(c => !(c in jt.idx) || fmt(nw[jt.idx[c]]) === '' ||
                                                 !isNaN(Number(fmt(nw[jt.idx[c]]).replace(/[$,]/g, ''))));
        if (okStatus && okDate && okNums) {
          changes.push({ kind: 'shift', row: row, header: 'whole row', label: label, oldRow: old, newRow: nw });
        } else {
          notTouched.push('row ' + row + ' | ' + label + ' | looks shifted, but the shifted result did not pass the status, date and number checks');
        }
      }
      return;
    }

    // 2. status word sitting in the Customer Name cell
    if (STATUS_LC[norm(name)] && norm(name) === norm(status)) {
      const contactName = byUid[uid];
      const estNames    = uniq(es.map(e => norm(e.name)).filter(Boolean));
      if (estNames.length === 1 && estNames[0] === norm(contactName)) {
        changes.push({ kind: 'name', row: row, col: jt.idx['Customer Name'] + 1, header: 'Customer Name', oldVal: name, newVal: contactName, label: 'job ' + jid });
      } else {
        notTouched.push('row ' + row + ' | job ' + jid + ' | status word in name, but contact ("' + contactName + '") and estimate ("' +
                        es.map(e => e.name).filter(Boolean).join('", "') + '") do not agree');
      }
    }
  });

  const counts = { uid: 0, name: 0, shift: 0 };
  changes.forEach(c => { counts[c.kind]++; });
  return { changes: changes, notTouched: notTouched, counts: counts, jt: jt, W: W };
}

function _jobsRepairRowText(vals, headers) {
  const parts = [];
  headers.forEach((h, c) => {
    const v = (vals[c] instanceof Date && !isNaN(vals[c].getTime())) ? vals[c].toISOString() : String(vals[c] == null ? '' : vals[c]).trim();
    if (h && v !== '') parts.push(h + '=' + (v.length > 50 ? v.substring(0, 50) + '...' : v));
  });
  return parts.join(' | ');
}

function repairJobsDryRun() {
  const p = _jobsRepairPlan();
  if (p.error) { Logger.log('REPAIR ERROR: ' + p.error); return; }

  const S = ['JOBS REPAIR DRY RUN (nothing was changed)  ' + new Date().toISOString(), '',
             'Would stamp a Contact UID:        ' + p.counts.uid,
             'Would fix a Customer Name cell:   ' + p.counts.name,
             'Would repair a shifted row:       ' + p.counts.shift,
             'Flagged but left alone for you:   ' + p.notTouched.length];
  Logger.log(S.join('\n'));

  const lines = kind => p.changes.filter(c => c.kind === kind);
  const U = lines('uid');
  if (U.length) Logger.log('CONTACT UID STAMPS\n' + U.map(c => '  row ' + c.row + ' | Contact UID | (blank) -> ' + c.newVal + ' | ' + c.label).join('\n'));
  const N = lines('name');
  if (N.length) Logger.log('CUSTOMER NAME FIXES\n' + N.map(c => '  row ' + c.row + ' | Customer Name | "' + c.oldVal + '" -> "' + c.newVal + '" | ' + c.label).join('\n'));
  lines('shift').forEach(c => {
    Logger.log('SHIFTED ROW REPAIR: row ' + c.row + ' | ' + c.label + '\n' +
               '  BEFORE: ' + _jobsRepairRowText(c.oldRow, p.jt.headers) + '\n' +
               '  AFTER:  ' + _jobsRepairRowText(c.newRow, p.jt.headers));
  });
  if (p.notTouched.length) Logger.log('LEFT ALONE (needs a person)\n  ' + p.notTouched.join('\n  '));
}

function repairJobsApply() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('Could not get the script lock. Nothing was changed. Try again in a minute.'); return; }
  try {
    const p = _jobsRepairPlan();
    if (p.error) { Logger.log('REPAIR ERROR: ' + p.error + '. Nothing was changed.'); return; }
    if (!p.changes.length) { Logger.log('Nothing to repair. Every row already passes. Nothing was changed.'); return; }

    const sh  = getSpreadsheet().getSheetByName(JOBS_TAB);
    const log = ['JOBS REPAIR APPLIED  ' + new Date().toISOString() + '  (save this log, it is your undo record)', ''];
    p.changes.forEach(c => {
      if (c.kind === 'shift') {
        sh.getRange(c.row, 1, 1, c.newRow.length).setValues([c.newRow]);
        log.push('SHIFTED ROW ' + c.row + ' | ' + c.label);
        log.push('  OLD: ' + _jobsRepairRowText(c.oldRow, p.jt.headers));
        log.push('  NEW: ' + _jobsRepairRowText(c.newRow, p.jt.headers));
      } else {
        sh.getRange(c.row, c.col).setValue(c.newVal);
        log.push('row ' + c.row + ' | ' + c.header + ' | "' + c.oldVal + '" -> "' + c.newVal + '" | ' + c.label);
      }
    });
    SpreadsheetApp.flush();
    log.push('');
    log.push('Changed: ' + p.counts.uid + ' Contact UIDs, ' + p.counts.name + ' Customer Names, ' + p.counts.shift + ' shifted rows.');
    log.push('Left alone: ' + p.notTouched.length + '. Run previewJobsRepair again to confirm.');
    Logger.log(log.join('\n'));
    if (p.notTouched.length) Logger.log('LEFT ALONE (needs a person)\n  ' + p.notTouched.join('\n  '));
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════
// JOBS COLUMN ORDER CHECK  (v2.19.7)
// READ ONLY. Changes no data and creates no sheets.
// The repair refused to touch job 949 because the Jobs tab's column order
// differs from the order the code writes in (JOB_HEADERS). saveJob() and
// updateStatus() write whole rows by position, so if the order differs, some
// values land under the wrong header. This function shows exactly where the
// two orders differ, then shows the rows that look out of place two ways:
//   sheet view = each value named by the header it sits under in the sheet
//   code view  = each value named by the position the code would write it to
// To run it: pick checkJobsColumnOrder in the Apps Script editor, press Run,
// then open Execution log.
// ══════════════════════════════════════════════════════════════
function checkJobsColumnOrder() {
  const sh = getSpreadsheet().getSheetByName(JOBS_TAB);
  if (!sh || sh.getLastRow() < 1) { Logger.log('Jobs sheet missing or empty'); return; }
  const vals = sh.getDataRange().getValues();
  const phys = vals[0].map(h => String(h).trim());
  const fmt  = v => (v instanceof Date && !isNaN(v.getTime())) ? v.toISOString() : String(v == null ? '' : v).trim();
  const clip = (s, n) => { s = String(s); return s.length > n ? s.substring(0, n) + '...' : s; };

  // ── Part 1: position by position ──
  const L = ['JOBS COLUMN ORDER CHECK (read only)  ' + new Date().toISOString(), '',
             'Sheet has ' + phys.length + ' columns. Code expects ' + JOB_HEADERS.length + '.', '',
             'position | code order | sheet order'];
  let diffs = 0;
  const top = Math.max(phys.length, JOB_HEADERS.length);
  for (let k = 0; k < top; k++) {
    const c = JOB_HEADERS[k] || '(none)', s = phys[k] || '(none)';
    const same = (c === s);
    if (!same) diffs++;
    L.push('  ' + (k + 1) + ' | ' + c + ' | ' + s + (same ? '' : '   <<< DIFFERENT'));
  }
  L.push('');
  L.push(diffs ? ('Positions that differ: ' + diffs) : 'Column order matches the code exactly.');
  const missing = JOB_HEADERS.filter(h => phys.indexOf(h) < 0);
  const extra   = phys.filter(h => h && JOB_HEADERS.indexOf(h) < 0);
  if (missing.length) L.push('In code but not in sheet: ' + missing.join(', '));
  if (extra.length)   L.push('In sheet but not in code: ' + extra.join(', '));
  Logger.log(L.join('\n'));

  // ── Part 2: rows that look out of place, shown both ways ──
  const KNOWN = ['Draft','Committed','Sent','Accepted','Deposit Received','Scheduled','In Progress',
                 'Ready for Invoice','Invoiced','Paid in Full','Lost'];
  const NUM   = ['Contract Price','Deposit Required','Deposit Paid','Final Payment Amount','Balance Due'];
  const pi    = h => phys.indexOf(h);
  const isDateLike = v => (v instanceof Date && !isNaN(v.getTime())) || (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v.trim()));

  function sheetView(r) {
    const p = []; phys.forEach((h, c) => { const v = fmt(r[c]); if (h && v !== '') p.push(h + '=' + clip(v, 40)); }); return p.join(' | ');
  }
  function codeView(r) {
    const p = []; JOB_HEADERS.forEach((h, c) => { const v = fmt(r[c]); if (v !== '') p.push(h + '=' + clip(v, 40)); }); return p.join(' | ');
  }

  const flagged = [], scheduledOk = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i], row = i + 1;
    const jid = pi('Job ID') >= 0 ? fmt(r[pi('Job ID')]) : '';
    if (!jid && !fmt(r[pi('Customer Name')])) continue;
    let bad = false;
    const st = pi('Status') >= 0 ? fmt(r[pi('Status')]) : '';
    if (st && KNOWN.indexOf(st) < 0) bad = true;
    if (pi('Status Changed At') >= 0) { const raw = r[pi('Status Changed At')]; if (raw && !isDateLike(raw)) bad = true; }
    NUM.forEach(h => { if (pi(h) >= 0) { const v = fmt(r[pi(h)]); if (v !== '' && isNaN(Number(v.replace(/[$,]/g, '')))) bad = true; } });
    if (bad) flagged.push({ row: row, jid: jid, vals: r });
    else if (pi('Scheduled Date') >= 0 && fmt(r[pi('Scheduled Date')]) !== '' && scheduledOk.length < 3) scheduledOk.push({ row: row, jid: jid, vals: r });
  }

  Logger.log('ROWS THAT LOOK OUT OF PLACE: ' + flagged.length + (flagged.length > 12 ? ' (showing first 12)' : ''));
  flagged.slice(0, 12).forEach(f => {
    Logger.log('row ' + f.row + ' | job ' + f.jid + '\n  sheet view: ' + sheetView(f.vals) + '\n  code view:  ' + codeView(f.vals));
  });
  if (scheduledOk.length) {
    Logger.log('FOR COMPARISON: scheduled jobs that look normal');
    scheduledOk.forEach(f => {
      Logger.log('row ' + f.row + ' | job ' + f.jid + '\n  sheet view: ' + sheetView(f.vals) + '\n  code view:  ' + codeView(f.vals));
    });
  }
}

// ══════════════════════════════════════════════════════════════
// JOBS LAYOUT PREVIEW  (v2.19.8)
// READ ONLY. Changes no data and creates no sheets.
// Background: the Jobs tab header row lists "Scheduled Time" near the end, but
// the code (JOB_HEADERS) writes it in the middle. So the Jobs tab holds two kinds
// of rows:
//   SHEET-ORDER rows  - written before the code changed; their cells line up with
//                       the headers as they stand in the sheet.
//   CODE-ORDER rows   - written since; their cells line up with JOB_HEADERS, not
//                       with the sheet headers.
// This function sorts every row into one of those kinds and shows what a clean-up
// would do: move the SHEET-ORDER rows into code order and make the header row
// match JOB_HEADERS. Nothing is written here.
// Test used: Final Payment Amount and Balance Due are always saved as numbers.
//   If those two cells are numbers under the SHEET headers      -> SHEET-ORDER
//   If those two cells are numbers at the CODE positions        -> CODE-ORDER
//   Both or neither                                             -> REVIEW BY HAND
// To run it: pick previewJobsLayout in the Apps Script editor, press Run, then
// open Execution log.
// ══════════════════════════════════════════════════════════════
function previewJobsLayout() {
  const sh = getSpreadsheet().getSheetByName(JOBS_TAB);
  if (!sh || sh.getLastRow() < 2) { Logger.log('Jobs sheet missing or empty'); return; }
  const vals = sh.getDataRange().getValues();
  const phys = vals[0].map(h => String(h).trim());
  const W    = JOB_HEADERS.length;
  const fmt  = v => (v instanceof Date && !isNaN(v.getTime())) ? v.toISOString() : String(v == null ? '' : v).trim();
  const clip = (s, n) => { s = String(s); return s.length > n ? s.substring(0, n) + '...' : s; };

  const missing = JOB_HEADERS.filter(h => phys.indexOf(h) < 0);
  const extra   = phys.filter(h => h && JOB_HEADERS.indexOf(h) < 0);
  if (phys.length !== W || missing.length || extra.length) {
    Logger.log('STOP: the Jobs headers are not the same set of ' + W + ' columns as the code.' +
               (missing.length ? '\nIn code but not in sheet: ' + missing.join(', ') : '') +
               (extra.length   ? '\nIn sheet but not in code: ' + extra.join(', ') : ''));
    return;
  }

  if (phys.length === W && JOB_HEADERS.every((h, k) => phys[k] === h)) { Logger.log('The Jobs header row already matches the code order. Nothing to classify.'); return; }

  // physIdx[k] = where the sheet keeps the column that the code calls position k
  const physIdx = JOB_HEADERS.map(h => phys.indexOf(h));
  const pFPA = phys.indexOf('Final Payment Amount'), pBD = phys.indexOf('Balance Due');
  const cFPA = JOB_HEADERS.indexOf('Final Payment Amount'), cBD = JOB_HEADERS.indexOf('Balance Due');
  const pST  = phys.indexOf('Scheduled Time'), cST = JOB_HEADERS.indexOf('Scheduled Time');
  const isNum = v => typeof v === 'number' && !isNaN(v);
  const timeOk = v => { const s = fmt(v); return s === '' || /^(morning|afternoon|full day|all day|am|pm)$/i.test(s) || /^\d{1,2}(:\d{2})?\s*(am|pm)?(\s*-\s*\d{1,2}(:\d{2})?\s*(am|pm)?)?$/i.test(s); };

  // A SHEET-ORDER row re-laid into code order (a pure re-arrangement, nothing lost)
  const toCodeOrder = r => JOB_HEADERS.map((h, k) => r[physIdx[k]]);
  const sheetView = r => { const p = []; phys.forEach((h, c) => { const v = fmt(r[c]); if (v !== '') p.push(h + '=' + clip(v, 30)); }); return p.join(' | '); };

  const rows = { SHEET: [], CODE: [], BOTH: [], NEITHER: [] };
  const runs = [];   // run-length map of classes down the sheet
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i], row = i + 1;
    if (!fmt(r[0]) && !fmt(r[3])) continue;
    const sheetSig = isNum(r[pFPA]) && isNum(r[pBD]);
    const codeSig  = isNum(r[cFPA]) && isNum(r[cBD]);
    const cls = sheetSig && codeSig ? 'BOTH' : sheetSig ? 'SHEET' : codeSig ? 'CODE' : 'NEITHER';
    rows[cls].push({ row: row, jid: fmt(r[0]), status: fmt(r[JOB_HEADERS.indexOf('Status')]), vals: r });
    const last = runs[runs.length - 1];
    if (last && last.cls === cls && last.end === row - 1) last.end = row; else runs.push({ cls: cls, start: row, end: row });
  }
  const total = rows.SHEET.length + rows.CODE.length + rows.BOTH.length + rows.NEITHER.length;

  Logger.log([
    'JOBS LAYOUT PREVIEW (read only)  ' + new Date().toISOString(), '',
    'Job rows examined: ' + total,
    '  SHEET-ORDER rows (would be moved into code order): ' + rows.SHEET.length,
    '  CODE-ORDER rows (already right for the code, would not move): ' + rows.CODE.length,
    '  REVIEW BY HAND - could be either: ' + rows.BOTH.length,
    '  REVIEW BY HAND - neither fits: ' + rows.NEITHER.length, '',
    'After the clean-up the header row would read, in order:',
    '  ' + JOB_HEADERS.join(' | ')
  ].join('\n'));

  // Where each kind sits down the sheet
  const R = ['ROW MAP (sheet row numbers, in order):'];
  runs.slice(0, 60).forEach(x => R.push('  rows ' + x.start + (x.end > x.start ? '-' + x.end : '') + ': ' + x.cls + ' (' + (x.end - x.start + 1) + ')'));
  if (runs.length > 60) R.push('  ... ' + (runs.length - 60) + ' more runs not shown');
  Logger.log(R.join('\n'));

  // Rows nobody can classify by the number test
  ['BOTH', 'NEITHER'].forEach(k => {
    if (!rows[k].length) return;
    const D = ['REVIEW BY HAND (' + k + '): ' + rows[k].length + (rows[k].length > 25 ? ' (showing first 25)' : '')];
    rows[k].slice(0, 25).forEach(f => D.push('  row ' + f.row + ' | job ' + f.jid + ' | ' + f.status + '\n    sheet view: ' + sheetView(f.vals)));
    Logger.log(D.join('\n'));
  });

  // Sanity checks on the Scheduled Time cell, which is the one that moves furthest
  const oddLegacy = rows.SHEET.filter(f => !timeOk(f.vals[pST]));
  const oddCode   = rows.CODE.filter(f => !timeOk(f.vals[cST]));
  Logger.log('SCHEDULED TIME CHECK\n' +
    '  SHEET-ORDER rows whose Scheduled Time cell looks unusual: ' + oddLegacy.length +
    (oddLegacy.length ? '\n' + oddLegacy.slice(0, 15).map(f => '    row ' + f.row + ' | job ' + f.jid + ' | "' + clip(fmt(f.vals[pST]), 40) + '"').join('\n') : '') + '\n' +
    '  CODE-ORDER rows whose Scheduled Time cell looks unusual: ' + oddCode.length +
    (oddCode.length ? '\n' + oddCode.slice(0, 15).map(f => '    row ' + f.row + ' | job ' + f.jid + ' | "' + clip(fmt(f.vals[cST]), 40) + '"').join('\n') : ''));

  // Show the move on real rows: before (as stored) and after (in code order)
  const pick = rows.SHEET.filter(f => fmt(f.vals[pST]) !== '' || fmt(f.vals[phys.indexOf('Schedule Notes')]) !== '' || (fmt(f.vals[pFPA]) !== '' && Number(f.vals[pFPA]) !== 0)).slice(0, 3);
  if (pick.length) {
    const E = ['EXAMPLE OF THE MOVE (first sheet-order rows that carry schedule or payment values):'];
    pick.forEach(f => {
      const after = toCodeOrder(f.vals);
      const p = []; JOB_HEADERS.forEach((h, k) => { const v = fmt(after[k]); if (v !== '') p.push(h + '=' + clip(v, 30)); });
      E.push('  row ' + f.row + ' | job ' + f.jid + '\n    now:   ' + sheetView(f.vals) + '\n    after: ' + p.join(' | '));
    });
    Logger.log(E.join('\n'));
  }
}

// ══════════════════════════════════════════════════════════════
// JOBS COLUMN ORDER FIX  (v2.19.9)
// Makes the Jobs tab header row match the order the code writes in (JOB_HEADERS)
// and moves the rows that were saved in the old sheet order into the code order.
//
// What it changes, and nothing else:
//   1. Rows classified SHEET-ORDER by previewJobsLayout(): only the cells in the
//      columns whose order differs are rearranged. No value is dropped or edited.
//   2. A row whose cells are all one column left of where they belong (job 949):
//      the missing Contact UID is inserted from its estimate, then the row is put
//      in code order. It is skipped if the estimate does not give exactly one
//      known contact, or if the result does not pass the status/date/number checks.
//   3. The header cells whose names differ are rewritten to the code's names.
// Rows that fit neither pattern (the "Committed" placeholders) are left alone.
//
//   layoutFixDryRun()  : prints exactly what would change. Changes nothing.
//   layoutFixApply()   : copies the Jobs tab to a backup tab first, then makes the
//                        changes, then re-reads the sheet to verify. Prints every
//                        old and new value. Save that log, it is your undo record.
//                        Safe to run twice: once the header matches, it does nothing.
// Run it when nobody is using the app. It takes the script lock while it works.
// ══════════════════════════════════════════════════════════════
function _jobsLayoutPlan() {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(JOBS_TAB);
  if (!sh || sh.getLastRow() < 2) return { error: 'Jobs sheet missing or empty' };
  const vals = sh.getDataRange().getValues();
  const phys = vals[0].map(h => String(h).trim());
  const W    = JOB_HEADERS.length;
  if (phys.length === W && JOB_HEADERS.every((h, k) => phys[k] === h)) return { alreadyDone: true };

  const missing = JOB_HEADERS.filter(h => phys.indexOf(h) < 0);
  const extra   = phys.filter(h => h && JOB_HEADERS.indexOf(h) < 0);
  if (phys.length !== W || missing.length || extra.length) {
    return { error: 'The Jobs headers are not the same set of ' + W + ' columns as the code.' +
                    (missing.length ? ' In code but not in sheet: ' + missing.join(', ') + '.' : '') +
                    (extra.length   ? ' In sheet but not in code: ' + extra.join(', ') + '.' : '') };
  }

  const norm  = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  const fmt   = v => (v instanceof Date && !isNaN(v.getTime())) ? v.toISOString() : String(v == null ? '' : v).trim();
  const uniq  = a => a.filter((x, i) => a.indexOf(x) === i);
  const isNum = v => typeof v === 'number' && !isNaN(v);
  const isDateLike = v => (v instanceof Date && !isNaN(v.getTime())) || (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v.trim()));
  const KNOWN = ['Draft','Committed','Sent','Accepted','Deposit Received','Scheduled','In Progress','Ready for Invoice','Invoiced','Paid in Full','Lost'];
  const STATUS_LC = {}; KNOWN.forEach(s => { STATUS_LC[s.toLowerCase()] = true; });
  const NUM = ['Contract Price','Deposit Required','Deposit Paid','Final Payment Amount','Balance Due'];

  const physIdx = JOB_HEADERS.map(h => phys.indexOf(h));          // sheet position of each code column
  const pi = h => phys.indexOf(h);
  const pFPA = pi('Final Payment Amount'), pBD = pi('Balance Due');
  const cFPA = JOB_HEADERS.indexOf('Final Payment Amount'), cBD = JOB_HEADERS.indexOf('Balance Due');
  const diffs = []; for (let k = 0; k < W; k++) if (phys[k] !== JOB_HEADERS[k]) diffs.push(k);
  const first = diffs[0], last = diffs[diffs.length - 1];
  const toCodeOrder = r => JOB_HEADERS.map((h, k) => r[physIdx[k]]);

  // Contacts and Estimates, read by header name (only needed for the shifted row)
  const rd = tab => {
    const s = ss.getSheetByName(tab); if (!s || s.getLastRow() < 1) return null;
    const v = s.getDataRange().getValues(), idx = {};
    v[0].forEach((h, i) => { const k = String(h).trim(); if (k && !(k in idx)) idx[k] = i; });
    return { idx: idx, rows: v.slice(1) };
  };
  const cell = (t, r, h) => (h in t.idx) ? fmt(r[t.idx[h]]) : '';
  const ct = rd(CONTACTS_TAB), et = rd(ESTIMATES_TAB);
  const byUid = Object.create(null);
  if (ct) ct.rows.forEach(r => { const u = cell(ct, r, 'UID'); if (u) byUid[u] = cell(ct, r, 'Name'); });

  const changes = [], leftAlone = [];
  let nSheet = 0, nCode = 0;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i], row = i + 1;
    const jid = fmt(r[0]), name = fmt(r[pi('Customer Name')]);
    if (!jid && !name) continue;
    const status = fmt(r[pi('Status')]);
    const label = 'job ' + jid + (name ? ' "' + name + '"' : '');
    const sheetSig = isNum(r[pFPA]) && isNum(r[pBD]);
    const codeSig  = isNum(r[cFPA]) && isNum(r[cBD]);

    if (sheetSig && !codeSig) {
      const nw = toCodeOrder(r);
      if (isNum(nw[cFPA]) && isNum(nw[cBD])) {
        changes.push({ kind: 'move', row: row, jid: jid, status: status, label: label, oldRow: r.slice(0, W), newRow: nw }); nSheet++;
      } else {
        leftAlone.push('row ' + row + ' | ' + label + ' | looked like sheet order but the moved result did not pass the number check');
      }
      continue;
    }
    if (codeSig && !sheetSig) { nCode++; continue; }
    if (codeSig && sheetSig) { leftAlone.push('row ' + row + ' | ' + label + ' | could be either order, a person should look'); continue; }

    // Fits neither: is it the "whole row one column left" pattern?
    const uidCell = r[pi('Contact UID')];
    const looksShifted = !!fmt(uidCell) && !(fmt(uidCell) in byUid) && !!STATUS_LC[norm(name)] &&
                         isDateLike(r[pi('Status')]) &&
                         !!fmt(r[pi('Status Changed At')]) && !isDateLike(r[pi('Status Changed At')]);
    if (!looksShifted) {
      leftAlone.push('row ' + row + ' | ' + label + ' | ' + (status || '(no status)') + ' | fits neither order, left as it is');
      continue;
    }
    if (!et) { leftAlone.push('row ' + row + ' | ' + label + ' | looks shifted, but the Estimates sheet could not be read'); continue; }
    const estId = fmt(r[pi('Estimate ID')]);
    const estUids = uniq(et.rows.filter(er => estId && cell(et, er, 'Estimate ID') === estId)
                                .map(er => cell(et, er, 'Contact UID')).filter(u => u && (u in byUid)));
    const old = r.slice(0, W);
    if (estUids.length !== 1) { leftAlone.push('row ' + row + ' | ' + label + ' | looks shifted, but the estimate does not give exactly one known contact'); continue; }
    if (fmt(old[W - 1]) !== '') { leftAlone.push('row ' + row + ' | ' + label + ' | looks shifted, but the last cell is not blank, so shifting would lose data'); continue; }
    const sh1 = old.slice(0, 2).concat([estUids[0]], old.slice(2, W - 1));      // still in sheet order
    const okStatus = KNOWN.indexOf(fmt(sh1[pi('Status')])) >= 0;
    const okDate   = isDateLike(sh1[pi('Status Changed At')]);
    const okNums   = NUM.every(c => fmt(sh1[pi(c)]) === '' || !isNaN(Number(fmt(sh1[pi(c)]).replace(/[$,]/g, ''))));
    const okSig    = isNum(sh1[pFPA]) && isNum(sh1[pBD]);
    const nw = toCodeOrder(sh1);
    if (okStatus && okDate && okNums && okSig && isNum(nw[cFPA]) && isNum(nw[cBD])) {
      changes.push({ kind: 'shift', row: row, jid: jid, status: fmt(sh1[pi('Status')]), label: label, oldRow: old, newRow: nw, uid: estUids[0], uidName: byUid[estUids[0]] });
    } else {
      leftAlone.push('row ' + row + ' | ' + label + ' | looks shifted, but the shifted result did not pass the status, date and number checks');
    }
  }
  return { changes: changes, leftAlone: leftAlone, phys: phys, first: first, last: last, diffs: diffs,
           counts: { move: changes.filter(c => c.kind === 'move').length, shift: changes.filter(c => c.kind === 'shift').length, code: nCode } };
}

function _layoutCells(arr, from, to) {
  const out = [];
  for (let k = from; k <= to; k++) {
    const v = arr[k];
    const s = (v instanceof Date && !isNaN(v.getTime())) ? v.toISOString() : String(v == null ? '' : v).trim();
    out.push(s.length > 22 ? s.substring(0, 22) + '...' : s);
  }
  return '[' + out.join(' | ') + ']';
}
function _layoutRowText(vals, headers) {
  const p = [];
  headers.forEach((h, c) => {
    const v = (vals[c] instanceof Date && !isNaN(vals[c].getTime())) ? vals[c].toISOString() : String(vals[c] == null ? '' : vals[c]).trim();
    if (h && v !== '') p.push(h + '=' + (v.length > 50 ? v.substring(0, 50) + '...' : v));
  });
  return p.join(' | ');
}

function layoutFixDryRun() {
  const p = _jobsLayoutPlan();
  if (p.error) { Logger.log('LAYOUT FIX ERROR: ' + p.error); return; }
  if (p.alreadyDone) { Logger.log('The Jobs header row already matches the code order. Nothing to do.'); return; }

  Logger.log([
    'JOBS COLUMN ORDER FIX, DRY RUN (nothing was changed)  ' + new Date().toISOString(), '',
    'Rows that would be rearranged (sheet order -> code order): ' + p.counts.move,
    'Rows that would get a missing Contact UID inserted, then rearranged: ' + p.counts.shift,
    'Rows already in code order, not touched: ' + p.counts.code,
    'Rows left alone: ' + p.leftAlone.length, '',
    'Header cells that would change (columns ' + (p.first + 1) + ' to ' + (p.last + 1) + '):',
    '  now:   ' + _layoutCells(p.phys, p.first, p.last),
    '  after: ' + _layoutCells(JOB_HEADERS, p.first, p.last)
  ].join('\n'));

  const mv = p.changes.filter(c => c.kind === 'move');
  for (let i = 0; i < mv.length; i += 12) {
    Logger.log('REARRANGED ROWS (cells in columns ' + (p.first + 1) + ' to ' + (p.last + 1) + ', before then after)\n' +
      mv.slice(i, i + 12).map(c => '  row ' + c.row + ' | ' + c.label + ' | ' + c.status + '\n    before: ' + _layoutCells(c.oldRow, p.first, p.last) +
                                   '\n    after:  ' + _layoutCells(c.newRow, p.first, p.last)).join('\n'));
  }
  p.changes.filter(c => c.kind === 'shift').forEach(c => {
    Logger.log('SHIFTED ROW: row ' + c.row + ' | ' + c.label + '\n' +
      '  Contact UID to insert: ' + c.uid + ' (contact "' + c.uidName + '")\n' +
      '  BEFORE (named by the current headers): ' + _layoutRowText(c.oldRow, p.phys) + '\n' +
      '  AFTER  (named by the new headers):     ' + _layoutRowText(c.newRow, JOB_HEADERS));
  });
  if (p.leftAlone.length) Logger.log('LEFT ALONE\n  ' + p.leftAlone.join('\n  '));
}

function layoutFixApply() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('Could not get the script lock. Nothing was changed. Try again in a minute.'); return; }
  try {
    const p = _jobsLayoutPlan();
    if (p.error) { Logger.log('LAYOUT FIX ERROR: ' + p.error + '. Nothing was changed.'); return; }
    if (p.alreadyDone) { Logger.log('The Jobs header row already matches the code order. Nothing was changed.'); return; }

    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(JOBS_TAB);
    const W  = JOB_HEADERS.length;

    // 1. Backup copy of the whole Jobs tab, verified before anything is written
    const stamp = new Date().toISOString().replace(/[-:T]/g, '').substring(0, 12);
    let bname = 'Jobs_backup_before_order_fix_' + stamp, n = 2;
    while (ss.getSheetByName(bname)) { bname = 'Jobs_backup_before_order_fix_' + stamp + '_' + n; n++; }
    const bk = sh.copyTo(ss);
    bk.setName(bname);
    SpreadsheetApp.flush();
    if (bk.getLastRow() !== sh.getLastRow() || bk.getLastColumn() !== sh.getLastColumn()) {
      Logger.log('The backup copy does not match the Jobs tab size. Nothing was changed. Backup tab: ' + bname);
      return;
    }
    const log = ['JOBS COLUMN ORDER FIX APPLIED  ' + new Date().toISOString() + '  (save this log, it is your undo record)',
                 'Backup of the Jobs tab before any change: tab "' + bname + '"', ''];

    // 2. Rows first, header last. A rerun after a stop is safe.
    p.changes.forEach(c => {
      if (c.kind === 'shift') sh.getRange(c.row, 1, 1, W).setValues([c.newRow]);
      else                    sh.getRange(c.row, p.first + 1, 1, p.last - p.first + 1).setValues([c.newRow.slice(p.first, p.last + 1)]);
    });
    sh.getRange(1, p.first + 1, 1, p.last - p.first + 1).setValues([JOB_HEADERS.slice(p.first, p.last + 1)]);
    SpreadsheetApp.flush();

    // 3. Undo record
    log.push('HEADER CELLS, columns ' + (p.first + 1) + ' to ' + (p.last + 1));
    log.push('  OLD: ' + _layoutCells(p.phys, p.first, p.last));
    log.push('  NEW: ' + _layoutCells(JOB_HEADERS, p.first, p.last));
    Logger.log(log.join('\n'));
    for (let i = 0; i < p.changes.length; i += 10) {
      Logger.log(p.changes.slice(i, i + 10).map(c => {
        if (c.kind === 'shift') {
          return 'SHIFTED ROW ' + c.row + ' | ' + c.label + ' | inserted Contact UID ' + c.uid + '\n  OLD ROW: ' + _layoutCells(c.oldRow, 0, W - 1) +
                 '\n  NEW ROW: ' + _layoutCells(c.newRow, 0, W - 1);
        }
        return 'row ' + c.row + ' | ' + c.label + '\n  OLD columns ' + (p.first + 1) + '-' + (p.last + 1) + ': ' + _layoutCells(c.oldRow, p.first, p.last) +
               '\n  NEW columns ' + (p.first + 1) + '-' + (p.last + 1) + ': ' + _layoutCells(c.newRow, p.first, p.last);
      }).join('\n'));
    }

    // 4. Verify by reading the sheet back
    const after = sh.getDataRange().getValues();
    const hdrOk = JOB_HEADERS.every((h, k) => String(after[0][k]).trim() === h) && after[0].length === W;
    const isNum = v => typeof v === 'number' && !isNaN(v);
    const bad = [];
    p.changes.forEach(c => {
      const r = after[c.row - 1];
      const same = JOB_HEADERS.every((h, k) => {
        const a = r[k], b = c.newRow[k];
        const sa = (a instanceof Date) ? a.getTime() : String(a == null ? '' : a);
        const sb = (b instanceof Date) ? b.getTime() : String(b == null ? '' : b);
        return sa === sb;
      });
      if (!same || !(isNum(r[JOB_HEADERS.indexOf('Final Payment Amount')]) && isNum(r[JOB_HEADERS.indexOf('Balance Due')]))) bad.push(c.row);
    });
    Logger.log([
      '', 'VERIFY',
      '  Header row matches the code: ' + (hdrOk ? 'YES' : 'NO'),
      '  Rows rearranged: ' + p.counts.move + '   rows shifted and rearranged: ' + p.counts.shift,
      '  Rows that failed the read-back check: ' + (bad.length ? bad.join(', ') : 'none'),
      '  Rows left alone: ' + p.leftAlone.length,
      '  Next: run checkJobsColumnOrder. It should say the column order matches exactly.'
    ].join('\n'));
    if (p.leftAlone.length) Logger.log('LEFT ALONE\n  ' + p.leftAlone.join('\n  '));
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════
// Tracks which sheets have had their headers validated this execution.
// Prevents redundant getRange calls on every getOrCreateSheet invocation.
const _validatedSheets = {};

function getOrCreateSheet(tabName, headers) {
  const ss    = getSpreadsheet();
  let sheet   = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(headers);
    const hRange = sheet.getRange(1,1,1,headers.length);
    hRange.setBackground('#1a1a1a').setFontColor('#e8c547')
          .setFontWeight('bold').setFontSize(10).setFontFamily('Courier New');
    sheet.setFrozenRows(1);
    _validatedSheets[tabName] = true;
  } else if (!_validatedSheets[tabName]) {
    // First call this execution only — check for missing columns, then mark as done.
    // All subsequent calls for this tab skip the header read entirely.
    const existingHdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>String(h).trim());
    let appended = false;
    headers.forEach(h => {
      if (!existingHdrs.includes(h)) {
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(h)
          .setBackground('#1a1a1a').setFontColor('#e8c547')
          .setFontWeight('bold').setFontSize(10).setFontFamily('Courier New');
        existingHdrs.push(h);
        appended = true;
      }
    });
    if (appended) sheet.setFrozenRows(1);
    _validatedSheets[tabName] = true;
  }
  return sheet;
}

function findRow(sheet, id, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const col = sheet.getRange(1, colIndex+1, lastRow, 1).getValues();
  for (let i = 1; i < col.length; i++) {
    if (String(col[i][0]).trim() === String(id).trim()) return i+1;
  }
  return -1;
}

// v3.20.2 — Didz's stated rule: one job never spans multiple companies, so
// picking a company on any one lettered option (A/B/C...) of a job should
// apply to every option of that same job, not just the one being edited.
// This updates every matching Estimates row's Company cell directly rather
// than re-running saveJob() per sibling, which would re-trigger photo
// uploads/new-contact detection/etc. for records the user isn't touching.
function updateCompanyForJob(data) {
  const prefix = String((data && data.jobPrefix) || '').trim();
  if (!prefix) return { error: 'Missing jobPrefix' };
  const company = (data && data.company) || 'BWI';
  const estSheet = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const lastRow = estSheet.getLastRow();
  if (lastRow < 2) return { updated: 0 };
  const prefixCol  = ESTIMATE_HEADERS.indexOf('Job Prefix') + 1;
  const companyCol = ESTIMATE_HEADERS.indexOf('Company') + 1;
  const prefixVals = estSheet.getRange(2, prefixCol, lastRow - 1, 1).getValues();
  let updated = 0;
  for (let i = 0; i < prefixVals.length; i++) {
    if (String(prefixVals[i][0]).trim() === prefix) {
      estSheet.getRange(i + 2, companyCol).setValue(company);
      updated++;
    }
  }
  return { updated: updated };
}

// v2.12.0 — set the Email and/or Phone cell on specific estimate rows (by
// Estimate ID). A blank value is never written, so nothing is ever wiped; at
// least one of email / phone must be supplied.
function updateEstimateContactInfo(data) {
  const email = String((data && data.email) || '').trim();
  const phone = String((data && data.phone) || '').trim();
  const ids   = (data && data.estimateIds) || [];
  if (!email && !phone) return { error: 'Missing email or phone' };
  if (!ids.length) return { updated: 0 };
  const estSheet = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const lastRow  = estSheet.getLastRow();
  if (lastRow < 2) return { updated: 0 };
  const emailCol = ESTIMATE_HEADERS.indexOf('Email') + 1;
  const phoneCol = ESTIMATE_HEADERS.indexOf('Phone') + 1;
  const wanted = {};
  ids.forEach(id => { wanted[String(id).trim()] = true; });
  const idVals = estSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let updated = 0;
  for (let i = 0; i < idVals.length; i++) {
    if (wanted[String(idVals[i][0]).trim()]) {
      if (email) estSheet.getRange(i + 2, emailCol).setValue(email);
      if (phone) estSheet.getRange(i + 2, phoneCol).setValue(phone);
      updated++;
    }
  }
  return { updated: updated };
}

// v2.11.2 action name kept so an older front end keeps working.
function updateEstimateEmails(data) {
  const email = String((data && data.email) || '').trim();
  if (!email) return { error: 'Missing email' };
  return updateEstimateContactInfo({ estimateIds: data.estimateIds, email: email });
}

function getNextEstimateLetter(sheet, prefix) {
  if (!prefix) return 'A';
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return 'A';
  let maxIdx = -1;
  data.slice(1).forEach(row => {
    const id = String(row[0]||'');
    if (id.startsWith(prefix+'-')) {
      const parts = id.split('-');
      if (parts.length >= 3) {
        const idx = letters.indexOf(parts[2]);
        if (idx > maxIdx) maxIdx = idx;
      }
    }
  });
  return letters[maxIdx+1] || 'A';
}

function ensureJobFolder(jobId, custName, isNew) {
  const root     = getRootFolder();
  const prefix   = jobId.match(/^(\d+)/);
  const jobNum   = prefix ? prefix[1] : jobId.split('-')[0];
  const year     = new Date().getFullYear();
  const safeName = (custName||'Unknown').replace(/[:\\/*?"<>|]/g,'').trim();
  const folderName = jobNum+' - '+safeName+' - '+year;
  const existing = root.getFoldersByName(folderName);
  let folder;
  if (existing.hasNext()) {
    folder = existing.next();
    if (!isNew) return folder.getUrl();
  } else {
    folder = root.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  ['Archive','Estimates','Vendor Invoices','Work Order'].forEach(n => getOrCreateSubfolder(folder,n));
  const estPhotos = getOrCreateSubfolder(folder, 'Estimate Photos');
  estPhotos.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder.getUrl();
}

function getJobFolder(jobId, custName) {
  const root   = getRootFolder();
  const prefix = jobId.match(/^(\d+)/);
  const jobNum = prefix ? prefix[1] : jobId.split('-')[0];
  const year   = new Date().getFullYear();
  if (custName) {
    const safeName   = custName.replace(/[:\\/*?"<>|]/g,'').trim();
    const folderName = jobNum+' - '+safeName+' - '+year;
    const byName     = root.getFoldersByName(folderName);
    if (byName.hasNext()) return byName.next();
  }
  const allFolders = root.getFolders();
  while (allFolders.hasNext()) {
    const f    = allFolders.next();
    const name = f.getName();
    if (name.startsWith(jobNum+' - ') && name.endsWith(' - '+year)) return f;
  }
  const byId = root.getFoldersByName(jobId);
  if (byId.hasNext()) return byId.next();
  const byNum = root.getFoldersByName(jobNum);
  if (byNum.hasNext()) return byNum.next();
  return null;
}

function getOrCreateSubfolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

// ══════════════════════════════════════════════════════════════
// ACCESS CONTROL
// ══════════════════════════════════════════════════════════════
const ACCESS_TAB     = 'Access_Control';
const ACCESS_HEADERS = ['Email', 'Added At', 'Added By', 'Active'];

function getAllowedEmailsFromSheet() {
  try {
    const sheet = getOrCreateSheet(ACCESS_TAB, ACCESS_HEADERS);
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return { emails: [OWNER_EMAIL] };
    const emails = data.slice(1)
      .filter(row => row[0] && row[3] !== 'No')
      .map(row => String(row[0]).trim().toLowerCase());
    if (!emails.includes(OWNER_EMAIL.toLowerCase())) emails.unshift(OWNER_EMAIL.toLowerCase());
    return { emails };
  } catch(e) {
    return { emails: [OWNER_EMAIL], error: e.message };
  }
}

function verifyAccess(data) {
  try {
    const email = String(data.email||'').trim().toLowerCase();
    if (!email) return { allowed: false, reason: 'No email provided' };
    if (email === OWNER_EMAIL.toLowerCase()) return { allowed: true };
    const result = getAllowedEmailsFromSheet();
    const allowed = result.emails.includes(email);
    return { allowed, reason: allowed ? 'Approved' : 'Email not on access list' };
  } catch(e) {
    return { allowed: false, reason: e.message };
  }
}

function saveAllowedEmailsToSheet(data) {
  try {
    const sheet  = getOrCreateSheet(ACCESS_TAB, ACCESS_HEADERS);
    const emails = data.emails||[];
    const now    = new Date().toISOString();
    const addedBy = data.addedBy||'Admin';
    const existing = sheet.getDataRange().getValues();
    const existingEmails = existing.slice(1).map(r => String(r[0]).trim().toLowerCase());
    emails.forEach(email => {
      const e = String(email).trim().toLowerCase();
      if (!e || e === OWNER_EMAIL.toLowerCase()) return;
      if (!existingEmails.includes(e)) {
        sheet.appendRow([e, now, addedBy, 'Yes']);
        existingEmails.push(e);
      }
    });
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

function removeAllowedEmailFromSheet(data) {
  try {
    const email  = String(data.email||'').trim().toLowerCase();
    if (email === OWNER_EMAIL.toLowerCase()) return { error: 'Cannot remove owner' };
    const sheet  = getOrCreateSheet(ACCESS_TAB, ACCESS_HEADERS);
    const values = sheet.getDataRange().getValues();
    for (let i = values.length-1; i >= 1; i--) {
      if (String(values[i][0]).trim().toLowerCase() === email) {
        sheet.getRange(i+1, 4).setValue('No');
        return { success: true };
      }
    }
    return { error: 'Email not found' };
  } catch(e) {
    return { error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// SEND EMAIL  ← FIXED: forces HTML rendering, never falls back
// ══════════════════════════════════════════════════════════════
// v2.8.2 — same "the frontend has been calling this for a while and it
// never existed" problem as updateInvoice, found in the same audit. Unlike
// updateInvoice, this one didn't fail silently — sendCIFormEmail() on the
// frontend checks res.success and alerts the error, so "Email ST-124 to
// Customer" would have shown a visible "Unknown action" error every time
// anyone actually tried it. Still, it means this button has never worked.
function sendCIFormPDF(data) {
  try {
    const to = data.custEmail || '';
    if (!to) return { error: 'No customer email address provided' };
    const pdfBase64 = data.pdfBase64 || '';
    if (!pdfBase64) return { error: 'No PDF data provided' };
    const fileName    = data.fileName    || 'ST-124-CI-Form.pdf';
    const companyName = data.companyName || 'FencePro';
    const custName     = data.custName    || 'Valued Customer';

    const blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', fileName);
    const subject  = 'ST-124 Certificate of Capital Improvement — ' + companyName;
    const textBody = 'Hi ' + custName + ',\n\n' +
      'Attached is the pre-filled ST-124 Certificate of Capital Improvement for your fence project.\n' +
      'Please print, sign, and reply with a photo or scan.\n\n' +
      'Thank you,\n' + companyName;

    GmailApp.sendEmail(to, subject, textBody, {
      attachments: [blob],
      name:        companyName,
      replyTo:     Session.getActiveUser().getEmail(),
      noReply:     false
    });

    return { success: true, sentTo: to };
  } catch(e) {
    return { error: e.toString() };
  }
}

function sendProposalEmail(data) {
  try {
    const to       = data.to       || '';
    const subject  = data.subject  || 'Your Fence Estimate';
    const htmlBody = data.htmlBody || '';
    const textBody = data.textBody || ' ';
    const custName = data.custName || '';
    // v2.8.0 — "From" display name now follows the estimate's selected
    // company instead of always showing Black Wolf Industries LLC.
    // Front end already knows which company is active; falls back to
    // BWI's name only if it wasn't passed (older client, or a direct
    // API call that predates this field).
    const fromName = data.companyName || 'Black Wolf Industries LLC';

    if (!to)       return { error: 'No recipient email address provided' };
    if (!htmlBody) return { error: 'No HTML body provided — email not sent' };

    GmailApp.sendEmail(to, subject, textBody, {
      htmlBody: htmlBody,
      name:     fromName,
      replyTo:  Session.getActiveUser().getEmail(),
      noReply:  false
    });

    Logger.log('Email sent to: ' + to + ' — ' + subject);
    return { success: true, sentTo: to };
  } catch(e) {
    return { error: e.toString() };
  }
}

function createProposalDraft(data) {
  try {
    const to       = data.to       || '';
    const subject  = data.subject  || 'Your Fence Estimate';
    const htmlBody = data.htmlBody || '';
    const textBody = data.textBody || ' ';
    const fromName = data.companyName || 'Black Wolf Industries LLC';

    if (!to)       return { error: 'No recipient email address provided' };
    if (!htmlBody) return { error: 'No HTML body provided — draft not created' };

    const draft = GmailApp.createDraft(to, subject, textBody, {
      htmlBody: htmlBody,
      name:     fromName,
      replyTo:  Session.getActiveUser().getEmail()
    });
    let account = '';
    try { account = Session.getEffectiveUser().getEmail() || ''; } catch(e) {}
    Logger.log('Draft created for: ' + to + ' — ' + subject);
    return { success: true, draftId: draft.getId(), messageId: draft.getMessageId(), account: account };
  } catch(e) {
    return { error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// PROPOSAL TOKEN SYSTEM
// ══════════════════════════════════════════════════════════════
const TOKEN_HEADERS = [
  'Token', 'Estimate ID', 'Customer Name', 'Customer Email',
  'Created At', 'Expires At', 'Status', 'Response', 'Response At',
  'Company', 'IP Note', 'Price Type Chosen', 'Contract Price',
  'Estimate IDs', 'Proposal Mode',
  'First Opened At', 'View Count'
];

function generateToken() {
  const bytes = [];
  for (let i = 0; i < 16; i++) bytes.push(Math.floor(Math.random()*256));
  return bytes.map(b => b.toString(16).padStart(2,'0')).join('');
}

function generateProposalToken(data) {
  try {
    const sheet          = getOrCreateSheet(TOKENS_TAB, TOKEN_HEADERS);
    const estimateIds    = data.estimateIds||(data.estimateId?[data.estimateId]:[]);
    const primaryId      = estimateIds[0]||data.estimateId||'';
    const estimateIdsStr = estimateIds.join(',');
    const proposalMode   = data.proposalMode||'single';
    const scriptUrl      = ScriptApp.getService().getUrl();
    const now            = new Date();

    const wantedCompany = data.company||'BWI';
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const rowEstIds  = String(rows[i][13]||rows[i][1]||'').trim();
      const rowMode    = String(rows[i][14]||'single').trim();
      const rowStatus  = String(rows[i][6]||'').trim();
      const rowExpires = new Date(rows[i][5]);
      // v2.10.1 — also require the company to match before reusing an
      // existing token. Without this, changing a job's company and
      // generating "a new" link would silently hand back the OLD token
      // (and whatever company it was created with) as long as the same
      // estimate(s)/mode were requested again within the 30-day window —
      // the company on the request was being computed correctly but never
      // even reaching the part of this function that would act on it.
      const rowCompany = String(rows[i][9]||'BWI').trim();
      if (rowEstIds===estimateIdsStr && rowMode===proposalMode && rowCompany===wantedCompany &&
          rowStatus!=='Accepted' && rowStatus!=='Declined' && rowExpires>now) {
        const existingToken = String(rows[i][0]).trim();
        const url = scriptUrl+'?token='+existingToken;
        return { success:true, token:existingToken, url, expires:rows[i][5], reused:true };
      }
    }

    const token   = generateToken();
    const expires = new Date(now.getTime()+30*24*60*60*1000);
    // v2.10.1 — this array had one extra blank placeholder before the
    // company value, shifting it from column 9 ("Company", where
    // validateToken() and everything downstream actually reads it) into
    // column 10 ("IP Note") instead — leaving the real Company cell blank
    // on every single proposal token ever created, for every company.
    // getCompanyRecord() treats a blank code as "use BWI", which is why
    // every hosted proposal page, PDF, and owner notification email has
    // always shown Black Wolf regardless of which company was actually
    // selected. Match TOKEN_HEADERS exactly: Token, Estimate ID, Customer
    // Name, Customer Email, Created At, Expires At, Status, Response,
    // Response At, Company, IP Note, Price Type Chosen, Contract Price,
    // Estimate IDs, Proposal Mode.
    sheet.appendRow([
      token, primaryId, data.custName||'', data.custEmail||'',
      now.toISOString(), expires.toISOString(), 'Pending',
      '', '', wantedCompany, '', '', '', estimateIdsStr, proposalMode
    ]);
    const url = scriptUrl+'?token='+token;
    return { success:true, token, url, expires:expires.toISOString(), reused:false };
  } catch(e) {
    return { error: e.toString() };
  }
}

function validateToken(token) {
  const sheet = getOrCreateSheet(TOKENS_TAB, TOKEN_HEADERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(token).trim()) {
      const expires = new Date(data[i][5]);
      const status  = String(data[i][6]).trim();
      if (new Date() > expires) return { error:'expired', row:i+1 };
      if (status==='Accepted'||status==='Declined') return { error:'used', status, row:i+1 };
      return {
        token:data[i][0], estimateId:data[i][1], custName:data[i][2], custEmail:data[i][3],
        createdAt:data[i][4], expiresAt:data[i][5], status:data[i][6], company:data[i][9],
        estimateIds:data[i][13]?String(data[i][13]).split(',').map(s=>s.trim()).filter(Boolean):[data[i][1]],
        proposalMode:data[i][14]||'single', row:i+1
      };
    }
  }
  return null;
}

function recordProposalResponse(data) {
  try {
    const tokenData = validateToken(data.token);
    if (!tokenData) return { error:'Invalid token' };
    if (tokenData.error) return { error:tokenData.error };

    const tokenSheet    = getOrCreateSheet(TOKENS_TAB, TOKEN_HEADERS);
    const now           = new Date().toISOString();
    const response      = data.response;
    const priceType     = data.priceType||'';
    const contractPrice = Number(data.contractPrice)||0;

    tokenSheet.getRange(tokenData.row,7).setValue(response);
    tokenSheet.getRange(tokenData.row,8).setValue(response);
    tokenSheet.getRange(tokenData.row,9).setValue(now);
    const ipNote = [
      data.ip        ? 'IP: '+data.ip : '',
      data.timestamp ? 'Time: '+data.timestamp : '',
      data.userAgent ? 'Agent: '+String(data.userAgent).substring(0,120) : ''
    ].filter(Boolean).join(' | ');
    tokenSheet.getRange(tokenData.row,11).setValue(ipNote);
    tokenSheet.getRange(tokenData.row,12).setValue(priceType);
    tokenSheet.getRange(tokenData.row,13).setValue(contractPrice);

    const depositRequired = response==='Accepted'&&contractPrice>0 ? Math.round(contractPrice*0.5*100)/100 : 0;
    const balanceDue      = response==='Accepted'&&contractPrice>0 ? Math.round((contractPrice-depositRequired)*100)/100 : 0;

    const jobId = String(tokenData.estimateId||'').split('-')[0];

    // The proposal token doesn't carry a Contact UID, but saveJob() already
    // created this Jobs row (with Contact UID set) when the estimate was
    // first drafted — look it up so updateStatus() doesn't blank it out.
    const jobSheetLookup    = getOrCreateSheet(JOBS_TAB, JOB_HEADERS);
    const existingJobRowIdx = findRow(jobSheetLookup, jobId, 0);
    let existingContactUID  = '';
    if (existingJobRowIdx > 0) {
      existingContactUID = String(
        jobSheetLookup.getRange(existingJobRowIdx, JOB_HEADERS.indexOf('Contact UID') + 1).getValue() || ''
      );
    }

    // Route through updateStatus() — same header-name-safe row builder the
    // in-app status changes use — instead of hand-building a positional
    // array that can silently drift out of sync with JOB_HEADERS.
    updateStatus({
      jobId,
      estimateId:         tokenData.estimateId || '',
      contactUID:         existingContactUID,
      custName:            tokenData.custName || '',
      status:              response==='Accepted' ? 'Accepted' : 'Lost',
      contractPrice,
      contractPriceType:  priceType,
      depositType:         '50pct',
      depositRequired,
      balanceDue,
      lostReason:          response==='Declined' ? 'Customer declined via proposal link' : '',
      notes:               data.message||''
    });

    sendNotificationEmail(tokenData, response, priceType, contractPrice, data.message||'');
    return { success:true, response };
  } catch(e) {
    return { error: e.toString() };
  }
}

function sendNotificationEmail(tokenData, response, priceType, contractPrice, message) {
  try {
    const isAccepted  = response === 'Accepted';
    const companyRec  = getCompanyRecord(tokenData.company);
    const emoji       = isAccepted ? '✅' : '❌';
    const priceLabel  = priceType === 'Cash' ? 'Cash / Check' : 'CC / Finance';
    const subject     = emoji + ' Proposal ' + response + ' — ' + tokenData.custName;
    const accentColor = isAccepted ? '#16A34A' : '#EF4444';
    const bgColor     = isAccepted ? '#F0FDF4' : '#FFF5F5';
    const now         = new Date().toLocaleString('en-US', {
      month:'long', day:'numeric', year:'numeric',
      hour:'numeric', minute:'2-digit', hour12:true
    });

    // ── Plain text fallback ──
    const plainLines = [
      'FencePro — Proposal ' + response,
      '================================',
      'Customer:   ' + tokenData.custName,
      'Email:      ' + (tokenData.custEmail || '—'),
      'Estimate:   ' + tokenData.estimateId,
      'Response:   ' + response,
    ];
    if (isAccepted) {
      plainLines.push('Price Type: ' + priceLabel);
      plainLines.push('Contract:   $' + Math.round(contractPrice).toLocaleString());
      plainLines.push('Deposit:    $' + Math.round(contractPrice * 0.5).toLocaleString() + ' (50%)');
    }
    plainLines.push('Time:       ' + now);
    if (message) plainLines.push('\nCustomer message:\n"' + message + '"');
    plainLines.push('\nLog in to FencePro to take action.');
    const plainText = plainLines.join('\n');

    // ── HTML email ──
    const htmlBody = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f4f5;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif">
<div style="max-width:520px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.10)">

  <!-- Header -->
  <div style="background:#1B2B4B;padding:24px 32px;display:flex;align-items:center;gap:14px">
    <div style="font-size:32px;line-height:1">${emoji}</div>
    <div>
      <div style="font-size:18px;font-weight:800;color:#fff;letter-spacing:-.3px">Proposal ${response}</div>
      <div style="font-size:12px;color:rgba(255,255,255,.55);margin-top:2px">${now}</div>
    </div>
  </div>

  <!-- Customer info -->
  <div style="padding:24px 32px;border-bottom:1px solid #f0f0f0">
    <div style="font-size:10px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">Customer</div>
    <div style="font-size:20px;font-weight:800;color:#1B2B4B;margin-bottom:4px">${escHtml(tokenData.custName)}</div>
    ${tokenData.custEmail ? `<div style="font-size:13px;color:#666"><a href="mailto:${escHtml(tokenData.custEmail)}" style="color:#2563EB;text-decoration:none">${escHtml(tokenData.custEmail)}</a></div>` : ''}
    <div style="font-size:12px;color:#999;margin-top:4px">Estimate: ${escHtml(tokenData.estimateId)}</div>
  </div>

  ${isAccepted ? `
  <!-- Accepted details -->
  <div style="padding:24px 32px;background:${bgColor};border-bottom:1px solid #e8f5e9">
    <div style="font-size:10px;font-weight:700;color:${accentColor};text-transform:uppercase;letter-spacing:.1em;margin-bottom:12px">Accepted Details</div>
    <div style="display:flex;justify-content:space-between;margin-bottom:8px">
      <span style="font-size:13px;color:#555">Price Type</span>
      <span style="font-size:13px;font-weight:700;color:#1B2B4B">${priceLabel}</span>
    </div>
    <div style="display:flex;justify-content:space-between;margin-bottom:8px">
      <span style="font-size:13px;color:#555">Contract Price</span>
      <span style="font-size:16px;font-weight:900;color:${accentColor};font-family:monospace">$${Math.round(contractPrice).toLocaleString()}</span>
    </div>
    <div style="display:flex;justify-content:space-between">
      <span style="font-size:13px;color:#555">Deposit Due (50%)</span>
      <span style="font-size:13px;font-weight:700;color:#1B2B4B;font-family:monospace">$${Math.round(contractPrice * 0.5).toLocaleString()}</span>
    </div>
  </div>` : `
  <!-- Declined -->
  <div style="padding:20px 32px;background:${bgColor};border-bottom:1px solid #fee2e2">
    <div style="font-size:13px;color:#666">The customer declined this proposal.${message ? ' See their message below.' : ''}</div>
  </div>`}

  ${message ? `
  <!-- Customer message -->
  <div style="padding:20px 32px;border-bottom:1px solid #f0f0f0">
    <div style="font-size:10px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px">Customer Message</div>
    <div style="font-size:13px;color:#444;font-style:italic;line-height:1.6;background:#f9f9f9;padding:12px;border-radius:6px;border-left:3px solid #ddd">"${escHtml(message)}"</div>
  </div>` : ''}

  <!-- CTA -->
  <div style="padding:24px 32px;text-align:center">
    <div style="font-size:13px;color:#666;margin-bottom:16px">Log in to FencePro to ${isAccepted ? 'record the deposit and schedule this job' : 'review and follow up'}.</div>
    <div style="font-size:11px;color:#aaa">${escHtml(companyRec.name)} &nbsp;·&nbsp; FencePro</div>
  </div>

</div>
</body>
</html>`;

    // Send to BOTH email addresses to make sure it arrives
    const recipients = NOTIFY_EMAIL;
    GmailApp.sendEmail(recipients, subject, plainText, {
      htmlBody: htmlBody,
      name: 'FencePro Notifications',
      noReply: false
    });
    Logger.log('Notification sent: ' + subject);
  } catch(e) {
    Logger.log('Email notification failed: ' + e.toString());
  }
}

function logProposalView(tokenRow) {
  try {
    const sheet        = getOrCreateSheet(TOKENS_TAB, TOKEN_HEADERS);
    const now          = new Date().toISOString();
    const firstOpenCol = 16;
    const viewCountCol = 17;
    const currentFirst = sheet.getRange(tokenRow, firstOpenCol).getValue();
    const currentCount = Number(sheet.getRange(tokenRow, viewCountCol).getValue())||0;
    if (!currentFirst) sheet.getRange(tokenRow, firstOpenCol).setValue(now);
    sheet.getRange(tokenRow, viewCountCol).setValue(currentCount+1);
  } catch(e) {
    Logger.log('logProposalView error: '+e.toString());
  }
}

function getProposalAudit(data) {
  try {
    const estimateId = String(data.estimateId||'').trim();
    if (!estimateId) return { tokens:[] };
    const sheet  = getOrCreateSheet(TOKENS_TAB, TOKEN_HEADERS);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { tokens:[] };
    const tokens = [];
    for (let i = 1; i < values.length; i++) {
      const row      = values[i];
      const primaryId = String(row[1]||'').trim();
      const allIds    = String(row[13]||'').split(',').map(s=>s.trim()).filter(Boolean);
      if (primaryId!==estimateId && !allIds.includes(estimateId)) continue;
      tokens.push({
        token:row[0], estimateId:primaryId, custName:row[2], custEmail:row[3],
        createdAt:row[4]?new Date(row[4]).toISOString():'',
        expiresAt:row[5]?new Date(row[5]).toISOString():'',
        status:String(row[6]||'Pending'), response:row[7],
        responseAt:row[8]?new Date(row[8]).toISOString():'',
        company:row[9], priceType:row[11], contractPrice:Number(row[12])||0,
        estimateIds:String(row[13]||''), mode:String(row[14]||'single'),
        firstOpenedAt:row[15]?new Date(row[15]).toISOString():'',
        viewCount:Number(row[16])||0
      });
    }
    tokens.sort((a,b) => new Date(b.createdAt)-new Date(a.createdAt));
    return { tokens };
  } catch(e) {
    return { tokens:[], error:e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// PROPOSAL PAGE RENDERER
// ══════════════════════════════════════════════════════════════
function serveProposalPage(token) {
  const tokenData = validateToken(token);
  const scriptUrl = ScriptApp.getService().getUrl();

  if (!tokenData) return HtmlService.createHtmlOutput(errorPage('This proposal link is invalid or does not exist.'));
  if (tokenData.error==='expired') return HtmlService.createHtmlOutput(errorPage('This proposal link has expired. Please contact us for a new one.'));
  if (tokenData.error==='used') return HtmlService.createHtmlOutput(errorPage('This proposal has already been '+tokenData.status+'. Please contact us if you have questions.'));

  const estSheet   = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const estData    = estSheet.getDataRange().getValues();
  const estHeaders = estData[0].map(h => String(h).trim());

  const proposalMode = tokenData.proposalMode||'single';
  const estimateIds  = tokenData.estimateIds||[tokenData.estimateId];

  const estimates = estimateIds.map(id => {
    const row = estData.slice(1).find(r => String(r[0]).trim()===String(id).trim());
    if (!row) return null;
    const est = {};
    estHeaders.forEach((h,i) => { est[h]=row[i]; });
    return est;
  }).filter(Boolean);

  if (estimates.length===0) return HtmlService.createHtmlOutput(errorPage('Estimate not found. Please contact us.'));

  logProposalView(tokenData.row);

  const first     = estimates[0];
  const companyRec = getCompanyRecord(tokenData.company);
  const company   = companyRec.name;
  const custName  = String(first['Customer Name']||'');
  const custFirst = custName.split(' ')[0];
  const jobAddr   = String(first['Job Address']||'');
  const jobCity   = String(first['Job City']||'');
  const jobState  = String(first['Job State']||'NY');
  const jobZip    = String(first['Zip']||'');
  const estId     = String(first['Estimate ID']||'').split('-')[0];
  const today     = new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'});
  const saved     = first['Saved At']?new Date(first['Saved At']).toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'}):today;
  const expires   = new Date(tokenData.expiresAt).toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'});
  const addrLine  = [jobCity,jobState,jobZip].filter(Boolean).join(', ');
  // Build dynamic deposit terms based on actual deposit settings saved with this estimate
  const depositSentence = buildDepositTermsSentence(first);
  const terms     = `${depositSentence}\n\nCustomer Obligations:\n1. Provide a clear area in the driveway for material to be delivered.\n2. Customer must fill out and sign NYS Capital Improvement form.\n3. We call 811 for all fence jobs. No exceptions.\n4. Customer is responsible for clearing fence line of any obstructions prior to fence installation.\n5. Any images attached to the proposal are for reference only and may be slightly different in person.`;

  let contentHtml = '';
  let jsData      = '';

  if (proposalMode==='single'||estimates.length===1) {
    const est       = estimates[0];
    const finalCC   = Number(est['Final Price CC'])||0;
    const finalCash = Number(est['Cash Price'])||0;
    const photoUrls = String(est['Photo URLs']||'').split(',').map(u=>u.trim()).filter(Boolean);
    contentHtml = buildSingleEstimateContent(est, finalCC, finalCash, photoUrls);
    const depType = String(est['Deposit Type']||'50pct');
    const depReq  = Number(est['Deposit Required']||0);
    jsData = `const ESTIMATES=[{finalCC:${finalCC},finalCash:${finalCash},letter:'',depositType:'${depType}',depositRequired:${depReq}}];
const MODE='single'; const TOKEN='${token}'; const SCRIPT_URL='${scriptUrl}';`;
  } else if (proposalMode==='options') {
    contentHtml = buildOptionsContent(estimates);
    const estsJs = estimates.map(est => {
      const letter = String(est['Estimate ID']||'').split('-')[2]||'';
      const dt=String(est['Deposit Type']||'50pct'),dr=Number(est['Deposit Required']||0); return `{finalCC:${Number(est['Final Price CC'])||0},finalCash:${Number(est['Cash Price'])||0},letter:'${escJs(letter)}',depositType:'${dt}',depositRequired:${dr}}`;
    }).join(',');
    jsData = `const ESTIMATES=[${estsJs}]; const MODE='options'; const TOKEN='${token}'; const SCRIPT_URL='${scriptUrl}';`;
  } else if (proposalMode==='combined') {
    contentHtml = buildCombinedContent(estimates);
    const combCC   = Math.ceil(estimates.reduce((s,e)=>s+(Number(e['Final Price CC'])||0),0)/25)*25;
    const combCash = Math.ceil(estimates.reduce((s,e)=>s+(Number(e['Cash Price'])||0),0)/25)*25;
    const depTypeC = String(estimates[0]?.['Deposit Type']||'50pct');
    const depReqC  = Number(estimates[0]?.['Deposit Required']||0);
    jsData = `const ESTIMATES=[{finalCC:${combCC},finalCash:${combCash},letter:'',depositType:'${depTypeC}',depositRequired:${depReqC}}];
const MODE='combined'; const TOKEN='${token}'; const SCRIPT_URL='${scriptUrl}';`;
  }

  const allPhotoUrls = proposalMode==='single'?(tokenData.estimateIds||[tokenData.estimateId]).reduce((acc,id)=>{
    const est = estimates.find(e=>String(e['Estimate ID']||'').trim()===String(id).trim());
    if (est) acc.push(...String(est['Photo URLs']||'').split(',').map(u=>u.trim()).filter(Boolean));
    return acc;
  },[]):[];

  const html = buildProposalShell(company,custName,custFirst,jobAddr,addrLine,estId,saved,expires,terms,contentHtml,jsData,allPhotoUrls,companyRec.addr,companyRec.phone);
  return HtmlService.createHtmlOutput(html)
    .setTitle(estId + '-Estimate')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function generateProposalPdf(token) {
  const tokenData = validateToken(token);
  if (!tokenData) return HtmlService.createHtmlOutput('<p>Invalid or expired token.</p>');

  const estSheet   = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const estData    = estSheet.getDataRange().getValues();
  const estHeaders = estData[0].map(h => String(h).trim());

  const proposalMode = tokenData.proposalMode||'single';
  const estimateIds  = tokenData.estimateIds||[tokenData.estimateId];

  const estimates = estimateIds.map(id => {
    const row = estData.slice(1).find(r => String(r[0]).trim()===String(id).trim());
    if (!row) return null;
    const est = {}; estHeaders.forEach((h,i)=>{ est[h]=row[i]; }); return est;
  }).filter(Boolean);

  if (estimates.length===0) return HtmlService.createHtmlOutput('<p>Estimate not found.</p>');

  const first      = estimates[0];
  const companyRec = getCompanyRecord(tokenData.company);
  const company    = companyRec.name;
  const custName   = String(first['Customer Name']||'');
  const custFirst  = custName.split(' ')[0];
  const jobAddr    = String(first['Job Address']||'');
  const jobCity    = String(first['Job City']||'');
  const jobState   = String(first['Job State']||'NY');
  const jobZip     = String(first['Zip']||'');
  const estId      = String(first['Estimate ID']||'').split('-')[0];
  const today      = new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'});
  const saved      = first['Saved At']?new Date(first['Saved At']).toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'}):today;
  const expires    = new Date(tokenData.expiresAt).toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'});
  const addrLine   = [jobCity,jobState,jobZip].filter(Boolean).join(', ');
  const depositSentence = buildDepositTermsSentence(first);
  const terms      = `${depositSentence}\n\nCustomer Obligations:\n1. Provide a clear area in the driveway for material to be delivered.\n2. Customer must fill out and sign NYS Capital Improvement form.\n3. We call 811 for all fence jobs. No exceptions.\n4. Customer is responsible for clearing fence line of any obstructions prior to fence installation.\n5. Any images attached to the proposal are for reference only and may be slightly different in person.`;

  let contentHtml = '';
  if (proposalMode==='single'||estimates.length===1) {
    const est=estimates[0];
    const finalCC=Number(est['Final Price CC'])||0;
    const finalCash=Number(est['Cash Price'])||0;
    const photoUrls=String(est['Photo URLs']||'').split(',').map(u=>u.trim()).filter(Boolean);
    contentHtml=buildSingleEstimateContent(est,finalCC,finalCash,photoUrls);
  } else if (proposalMode==='options') {
    contentHtml=buildOptionsContent(estimates);
  } else if (proposalMode==='combined') {
    contentHtml=buildCombinedContent(estimates);
  }

  const allPhotoUrls = proposalMode==='single'?(tokenData.estimateIds||[tokenData.estimateId]).reduce((acc,id)=>{
    const est=estimates.find(e=>String(e['Estimate ID']||'').trim()===String(id).trim());
    if (est) acc.push(...String(est['Photo URLs']||'').split(',').map(u=>u.trim()).filter(Boolean));
    return acc;
  },[]):[];

  // Serve clean print-ready HTML — window.print() works in a direct tab (not iframe)
  const pdfHtml = buildProposalShellPdf(company,custName,custFirst,jobAddr,addrLine,estId,saved,expires,terms,contentHtml,allPhotoUrls,companyRec.addr,companyRec.phone);
  return HtmlService.createHtmlOutput(pdfHtml)
    .setTitle(estId + '-Estimate')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildProposalShell(company,custName,custFirst,jobAddr,addrLine,estId,saved,expires,terms,contentHtml,jsData,photoUrls,companyAddr,companyPhone) {
  photoUrls = photoUrls||[];
  // v2.8.0 — companyAddr/companyPhone are new, optional trailing params;
  // fall back to BWI/TRF's shared office if a caller doesn't pass them.
  companyAddr  = companyAddr  || '169 Commack Rd #1009, Commack, NY 11725';
  companyPhone = companyPhone || '631-275-2020';
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Proposal — ${escJs(company)}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Helvetica Neue',Arial,sans-serif;background:#E8E8E8;color:#222;font-size:14px;line-height:1.5;min-height:100vh}
  .page{max-width:780px;margin:30px auto;background:#fff;padding:52px 60px;box-shadow:0 4px 32px rgba(0,0,0,.12);border-radius:4px}
  @media(max-width:640px){.page{margin:0;padding:32px 24px;border-radius:0;box-shadow:none}}
  .header{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:20px;margin-bottom:28px;border-bottom:3px solid #1B2B4B}
  .co-name{font-size:22px;font-weight:900;color:#1B2B4B;letter-spacing:-.4px;margin-bottom:4px}
  .co-sub{font-size:11px;color:#999;line-height:1.7}
  .bill-label{font-size:9px;font-weight:700;color:#999;text-transform:uppercase;letter-spacing:.12em;margin-bottom:5px;margin-top:14px}
  .cust-name{font-size:14px;font-weight:700;color:#1B2B4B}
  .cust-addr{font-size:12px;color:#555;line-height:1.5}
  .est-badge{font-size:34px;font-weight:900;color:#F26522;letter-spacing:-1px;line-height:1;margin-bottom:10px;text-align:right}
  .meta-table td{font-size:11px;padding:2px 0;vertical-align:top}
  .meta-table td:first-child{color:#1B2B4B;font-weight:700;padding-right:14px;text-align:right;white-space:nowrap}
  .meta-table td:last-child{color:#444;min-width:100px}
  .section-label{font-size:9px;font-weight:700;color:#1B2B4B;text-transform:uppercase;letter-spacing:.12em;margin-bottom:8px}
  .scope-box{background:#F7F8FA;border-left:3px solid #1B2B4B;border-radius:4px;padding:14px 18px;margin-bottom:20px;white-space:pre-wrap;font-size:12px;color:#333;line-height:1.7}
  table.line{width:100%;border-collapse:collapse;margin-bottom:0}
  table.line thead tr{background:#1B2B4B;color:#fff}
  table.line th{padding:11px 16px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em}
  table.line th:last-child{text-align:right}
  table.line td{padding:14px 16px;font-size:13px;border-bottom:2px solid #E2E5EA;vertical-align:top}
  table.line td:last-child{text-align:right;font-weight:800;color:#1B2B4B;font-size:14px;white-space:nowrap}
  .totals{display:flex;justify-content:flex-end;margin:16px 0 28px}
  .totals table{border-collapse:collapse;width:300px}
  .totals td{padding:5px 0;font-size:13px}
  .totals td:first-child{text-align:right;padding-right:16px;color:#555}
  .totals td:last-child{text-align:right;font-weight:700}
  .total-cash{font-size:20px;font-weight:900;color:#16A34A}
  .total-cc{font-size:13px;font-weight:600;color:#F26522;opacity:.85}
  .terms-box{background:#F7F8FA;border-left:3px solid #1B2B4B;border-radius:4px;padding:14px 18px;margin-bottom:24px}
  .terms-box p{font-size:12px;color:#444;line-height:1.8;white-space:pre-wrap}
  .proceed-box{background:#FFF8F4;border-left:3px solid #F26522;border-radius:4px;padding:14px 18px;margin-bottom:32px}
  .proceed-box p{font-size:12px;color:#444;line-height:1.7}
  .proceed-box a{color:#F26522;font-weight:700;text-decoration:none}
  .opt-tabs{display:flex;gap:0;border-bottom:2px solid #1B2B4B;margin-bottom:24px}
  .opt-tab{padding:10px 24px;font-size:13px;font-weight:700;cursor:pointer;border:none;background:transparent;color:#999;border-bottom:3px solid transparent;margin-bottom:-2px;font-family:inherit;transition:.15s}
  .opt-tab.active{color:#F26522;border-bottom-color:#F26522}
  .opt-panel{display:none}.opt-panel.active{display:block}
  .payment-step{margin-bottom:28px}
  .payment-step-title{font-size:16px;font-weight:800;color:#1B2B4B;margin-bottom:6px;text-align:center}
  .payment-step-sub{font-size:13px;color:#666;margin-bottom:18px;text-align:center}
  .price-cards{display:flex;gap:14px;justify-content:center;flex-wrap:wrap}
  .price-card{flex:1;min-width:200px;max-width:260px;border:2.5px solid #E2E5EA;border-radius:14px;padding:20px 18px;cursor:pointer;transition:all .2s;text-align:center;background:#fff}
  .price-card:hover{border-color:#1B2B4B;box-shadow:0 4px 20px rgba(27,43,75,.12);transform:translateY(-2px)}
  .price-card.selected-cc{border-color:#F26522;background:#FFF8F4;box-shadow:0 4px 20px rgba(242,101,34,.2)}
  .price-card.selected-cash{border-color:#16A34A;background:#F0FDF4;box-shadow:0 4px 20px rgba(22,163,74,.2)}
  .price-card .badge{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px}
  .price-card .amount{font-size:28px;font-weight:900;font-family:monospace;margin-bottom:4px;letter-spacing:-1px}
  .price-card .deposit{font-size:12px;color:#888;margin-bottom:10px}
  .price-card .check-ring{width:24px;height:24px;border-radius:50%;border:2px solid #ccc;display:flex;align-items:center;justify-content:center;margin:0 auto;font-size:14px;transition:.2s}
  .price-card.selected-cc .check-ring{background:#F26522;border-color:#F26522;color:#fff}
  .price-card.selected-cash .check-ring{background:#16A34A;border-color:#16A34A;color:#fff}
  .action-section{text-align:center;margin-bottom:32px}
  .action-title{font-size:16px;font-weight:800;color:#1B2B4B;margin-bottom:6px}
  .action-sub{font-size:13px;color:#666;margin-bottom:20px}
  .btn-group{display:flex;gap:16px;justify-content:center;flex-wrap:wrap}
  .btn-accept{display:inline-flex;align-items:center;gap:10px;padding:16px 40px;border-radius:10px;border:none;cursor:pointer;font-family:inherit;font-size:16px;font-weight:800;letter-spacing:-.2px;background:linear-gradient(135deg,#16A34A 0%,#15803D 100%);color:#fff;box-shadow:0 4px 20px rgba(22,163,74,.3);transition:all .2s;text-decoration:none;opacity:.4;pointer-events:none}
  .btn-accept.ready{opacity:1;pointer-events:auto}
  .btn-accept.ready:hover{transform:translateY(-2px);box-shadow:0 8px 28px rgba(22,163,74,.4)}
  .btn-decline{display:inline-flex;align-items:center;gap:10px;padding:16px 40px;border-radius:10px;border:2px solid #E2E5EA;cursor:pointer;font-family:inherit;font-size:16px;font-weight:800;letter-spacing:-.2px;background:#fff;color:#666;transition:all .2s;text-decoration:none}
  .btn-decline:hover{border-color:#EF4444;color:#EF4444;transform:translateY(-2px);box-shadow:0 4px 16px rgba(239,68,68,.15)}
  .message-box{margin-top:20px;display:none}
  .message-box textarea{width:100%;padding:12px;border:1px solid #E2E5EA;border-radius:8px;font-family:inherit;font-size:13px;color:#444;resize:vertical;min-height:80px;outline:none}
  .message-box textarea:focus{border-color:#1B2B4B}
  .message-box label{font-size:12px;font-weight:700;color:#555;margin-bottom:6px;display:block}
  .submit-btn{margin-top:10px;padding:10px 28px;border-radius:8px;border:none;cursor:pointer;font-family:inherit;font-size:14px;font-weight:700;color:#fff;background:#1B2B4B;transition:.15s}
  .submit-btn:hover{opacity:.85}
  .price-required-note{font-size:12px;color:#F26522;font-weight:600;margin-bottom:16px;display:none}
  .response-done{display:none;text-align:center;padding:32px;border-radius:12px;margin-bottom:32px}
  .response-done.accepted{background:#F0FDF4;border:2px solid #16A34A}
  .response-done.declined{background:#FFF5F5;border:2px solid #EF4444}
  .response-done .icon{font-size:48px;margin-bottom:12px}
  .response-done h2{font-size:20px;font-weight:800;margin-bottom:6px}
  .response-done p{font-size:13px;color:#666}
  .footer{display:flex;justify-content:space-between;align-items:center;padding-top:16px;border-top:1px solid #E2E5EA;font-size:11px;color:#aaa}
  /* ── Print toolbar ── */
  #print-toolbar{position:fixed;top:0;left:0;right:0;z-index:9999;background:#1B2B4B;padding:10px 20px;display:flex;align-items:center;gap:12px;box-shadow:0 2px 8px rgba(0,0,0,.25)}
  #print-toolbar .ptb-title{font-size:13px;font-weight:800;color:#fff}
  #print-toolbar .ptb-sub{font-size:11px;color:rgba(255,255,255,.5)}
  #print-toolbar .ptb-btn{margin-left:auto;padding:8px 20px;border-radius:7px;border:none;cursor:pointer;font-family:inherit;font-size:13px;font-weight:800;background:#F26522;color:#fff;letter-spacing:.02em;transition:.15s}
  #print-toolbar .ptb-btn:hover{background:#d4541a}
  body.has-toolbar .page{margin-top:64px}
  /* ── Print styles ── */
  @media print{
    #print-toolbar{display:none!important}
    body{background:#fff!important}
    body.has-toolbar .page{margin-top:0!important}
    .page{box-shadow:none!important;border-radius:0!important;margin:0!important;padding:40px 50px!important;max-width:100%!important}
    .btn-accept,.btn-decline,.action-section,.proceed-box,.payment-step,.price-cards,.btn-group,.message-box,.submit-btn,.price-required-note,.response-done{display:none!important}
    .opt-tabs{display:none!important}
    .opt-panel{display:block!important}
    a{text-decoration:none!important;color:inherit!important}
    @page{margin:0.5in}
  }
</style>
</head>
<body class="has-toolbar">
<div id="print-toolbar">
  <span class="ptb-title">&#128196; Proposal Preview</span>
  <span class="ptb-sub">This is exactly what your customer sees</span>
  <button class="ptb-btn" onclick="window.print()">&#8595; Save as PDF</button>
</div>
<div class="page">
  <div class="header">
    <div>
      <div class="co-name">${escHtml(company)}</div>
      <div class="co-sub">${escHtml(companyAddr)}<br>${escHtml(companyPhone)}</div>
      <div class="bill-label">Bill To</div>
      <div class="cust-name">${escHtml(custName)}</div>
      <div class="cust-addr">${escHtml(jobAddr)}${addrLine?'<br>'+escHtml(addrLine):''}</div>
    </div>
    <div style="text-align:right;flex-shrink:0;margin-left:40px">
      <div class="est-badge">ESTIMATE</div>
      <table class="meta-table" style="margin-left:auto">
        <tr><td>Job #</td><td>${escHtml(estId)}</td></tr>
        <tr><td>Date</td><td>${saved}</td></tr>
        <tr><td>Expires</td><td>${expires}</td></tr>
      </table>
    </div>
  </div>
  ${contentHtml}
  <div class="response-done accepted" id="resp_accepted">
    <div class="icon">✅</div>
    <h2 style="color:#16A34A">Proposal Accepted!</h2>
    <p>Thank you, ${escHtml(custFirst)}! We've received your acceptance and will be in touch shortly to arrange the deposit and schedule your installation.</p>
  </div>
  <div class="response-done declined" id="resp_declined">
    <div class="icon">❌</div>
    <h2 style="color:#EF4444">Proposal Declined</h2>
    <p>We've recorded your response. Please don't hesitate to reach out if you have any questions or would like to revisit the estimate.</p>
  </div>
  <div id="payment_step" style="display:none;margin-bottom:28px">
    <div class="payment-step-title">How would you like to pay?</div>
    <div class="payment-step-sub">Please select your payment method to complete your acceptance</div>
    <div class="price-cards">
      <div class="price-card" id="card_cc" onclick="selectPrice('CC')">
        <div class="badge" style="color:#F26522">?? Credit Card / Finance</div>
        <div class="amount" id="amt_cc" style="color:#F26522">$0</div>
        <div class="deposit" id="dep_cc">50% deposit: $0</div>
        <div class="check-ring" id="ring_cc">✓</div>
      </div>
      <div class="price-card" id="card_cash" onclick="selectPrice('Cash')">
        <div class="badge" style="color:#16A34A">?? Cash / Check</div>
        <div class="amount" id="amt_cash" style="color:#16A34A">$0</div>
        <div class="deposit" id="dep_cash">50% deposit: $0</div>
        <div class="check-ring" id="ring_cash">✓</div>
      </div>
    </div>
    <div style="text-align:center;margin-top:18px">
      <div class="price-required-note" id="price_required">⚠️ Please select a payment method above to continue</div>
      <button class="submit-btn" onclick="submitAcceptance()" style="background:#16A34A;font-size:15px;padding:14px 40px;border-radius:10px;margin-top:4px">✓ Confirm Acceptance</button>
      <br><button onclick="cancelAccept()" style="margin-top:10px;background:none;border:none;cursor:pointer;font-family:inherit;font-size:13px;color:#999;text-decoration:underline">Cancel</button>
    </div>
  </div>
  <div class="message-box" id="decline_box">
    <label>Let us know why (optional)</label>
    <textarea id="decline_msg" placeholder="Any feedback or questions for the team..."></textarea>
    <div style="display:flex;gap:10px;justify-content:center;margin-top:10px">
      <button class="submit-btn" onclick="submitDecline()" style="background:#EF4444">✕ Confirm Decline</button>
      <button onclick="cancelDecline()" style="padding:10px 20px;border-radius:8px;border:1px solid #E2E5EA;background:#fff;cursor:pointer;font-family:inherit;font-size:14px;color:#666">Cancel</button>
    </div>
  </div>
  <div class="action-section" id="action_section">
    <div class="action-title">Ready to proceed?</div>
    <div class="action-sub">Let us know your decision below</div>
    <div class="btn-group">
      <button class="btn-accept ready" id="btn_accept" onclick="showPaymentChoice()">✓ &nbsp;Accept Proposal</button>
      <button class="btn-decline" onclick="showDecline()">✕ &nbsp;Decline</button>
    </div>
  </div>
  ${photoUrls.length>0?`
  <div style="padding:0 0 24px">
    <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.08em;margin-bottom:12px">
      IMAGES &nbsp;<span style="font-weight:400;color:#aaa">For reference only — actual installation may vary slightly</span>
    </div>
    <div style="display:grid;grid-template-columns:${photoUrls.length===1?'1fr':photoUrls.length===2?'1fr 1fr':'repeat(2,1fr)'};gap:10px;">
      ${photoUrls.map(url=>{
        const src=url.includes('lh3.googleusercontent.com')?url:(()=>{const m=url.match(/[-\w]{25,}/);return m?'https://lh3.googleusercontent.com/d/'+m[0]:url;})();
        return '<img src="'+src+'" style="width:100%;height:auto;border-radius:8px;border:1px solid #E2E5EA;display:block;object-fit:cover" onerror="this.style.display=\'none\'">';
      }).join('')}
    </div>
  </div>`:''}
  <div class="footer">
    <span>© ${new Date().getFullYear()} ${escHtml(company)}</span>
    <span>This estimate is valid until ${expires}</span>
  </div>
</div>
<script>
${jsData}
let selectedPrice=0,selectedAmount=0,activeEstIdx=0;
function fmt(n){return '$'+Math.round(n).toLocaleString();}
function showPaymentChoice(){
  const est=ESTIMATES[MODE==='options'?activeEstIdx:0];
  document.getElementById('amt_cc').textContent=fmt(est.finalCC);
  const depPct = (est.depositType==='none') ? 0 : (est.depositType==='custom' ? (est.customDepositPct||50) : 50);
  const depAmt_cc   = est.finalCC   * (depPct/100);
  const depAmt_cash = est.finalCash * (depPct/100);
  document.getElementById('dep_cc').textContent   = depPct===0 ? 'No deposit required' : Math.round(depPct)+'% deposit: '+fmt(depAmt_cc);
  document.getElementById('amt_cash').textContent=fmt(est.finalCash);
  document.getElementById('dep_cash').textContent = depPct===0 ? 'No deposit required' : Math.round(depPct)+'% deposit: '+fmt(depAmt_cash);
  document.getElementById('action_section').style.display='none';
  document.getElementById('payment_step').style.display='block';
  document.getElementById('decline_box').style.display='none';
  window.scrollTo({top:document.getElementById('payment_step').offsetTop-20,behavior:'smooth'});
}
function cancelAccept(){selectedPrice='';selectedAmount=0;resetCards();document.getElementById('payment_step').style.display='none';document.getElementById('action_section').style.display='block';document.getElementById('price_required').style.display='none';}
function selectPrice(type){const est=ESTIMATES[MODE==='options'?activeEstIdx:0];selectedPrice=type;selectedAmount=type==='CC'?est.finalCC:est.finalCash;document.getElementById('card_cc').className='price-card'+(type==='CC'?' selected-cc':'');document.getElementById('card_cash').className='price-card'+(type==='Cash'?' selected-cash':'');document.getElementById('price_required').style.display='none';}
function resetCards(){document.getElementById('card_cc').className='price-card';document.getElementById('card_cash').className='price-card';}
async function submitAcceptance(){
  if(!selectedPrice){document.getElementById('price_required').style.display='block';return;}
  const btn=document.querySelector('#payment_step .submit-btn');btn.textContent='Sending...';btn.disabled=true;
  let ip='';try{const r=await fetch('https://api.ipify.org?format=json');ip=(await r.json()).ip||'';}catch(e){ip='unavailable';}
  try{
    const res=await fetch(SCRIPT_URL,{method:'POST',headers:{'Content-Type':'text/plain'},body:JSON.stringify({action:'recordProposalResponse',data:{token:TOKEN,response:'Accepted',priceType:selectedPrice,contractPrice:selectedAmount,message:'',ip,userAgent:navigator.userAgent,timestamp:new Date().toISOString()}})});
    const data=await res.json();
    if(data.success){document.getElementById('payment_step').style.display='none';document.getElementById('resp_accepted').style.display='block';window.scrollTo({top:0,behavior:'smooth'});}
    else{btn.textContent='Error — try again';btn.disabled=false;}
  }catch(e){btn.textContent='Error — try again';btn.disabled=false;}
}
function showDecline(){document.getElementById('action_section').style.display='none';document.getElementById('payment_step').style.display='none';document.getElementById('decline_box').style.display='block';}
function cancelDecline(){document.getElementById('decline_box').style.display='none';document.getElementById('action_section').style.display='block';}
async function submitDecline(){
  const btn=document.querySelector('#decline_box .submit-btn');btn.textContent='Sending...';btn.disabled=true;
  let ip='';try{const r=await fetch('https://api.ipify.org?format=json');ip=(await r.json()).ip||'';}catch(e){ip='unavailable';}
  try{
    const res=await fetch(SCRIPT_URL,{method:'POST',headers:{'Content-Type':'text/plain'},body:JSON.stringify({action:'recordProposalResponse',data:{token:TOKEN,response:'Declined',priceType:'',contractPrice:0,message:document.getElementById('decline_msg').value.trim(),ip,userAgent:navigator.userAgent,timestamp:new Date().toISOString()}})});
    const data=await res.json();
    if(data.success){document.getElementById('decline_box').style.display='none';document.getElementById('resp_declined').style.display='block';}
    else{btn.textContent='Error — try again';btn.disabled=false;}
  }catch(e){btn.textContent='Error — try again';btn.disabled=false;}
}
function switchOptTab(idx){activeEstIdx=idx;document.querySelectorAll('.opt-tab').forEach((t,i)=>t.classList.toggle('active',i===idx));document.querySelectorAll('.opt-panel').forEach((p,i)=>p.classList.toggle('active',i===idx));}
</script>
</body>
</html>`;
}

// ── Build dynamic deposit terms sentence based on actual deposit settings ──
function buildDepositTermsSentence(est) {
  // depositType: '50pct', 'custom', or 'none' — now stored in Estimates sheet col 59
  // depositRequired: dollar amount — stored in Estimates sheet col 60
  const depositType = String(est['Deposit Type'] || est['depositType'] || '50pct').trim();
  const depositReq  = Number(est['Deposit Required'] || est['depositRequired'] || 0);

  if (depositType === 'none') {
    return 'Full balance is due upon completion of the work.';
  }
  if (depositType === 'custom' && depositReq > 0) {
    return `A $${Math.round(depositReq).toLocaleString()} deposit is required to secure the job/materials/scheduling time. Balance due upon completion of the work.`;
  }
  // Default: 50%
  return '50% deposit is required to secure the job/materials/scheduling time. Balance due upon completion of the work.';
}

// ── Build dynamic proceed text based on deposit settings ──
function buildProceedText(est, company) {
  const depositType = String(est['Deposit Type'] || est['depositType'] || '50pct').trim();
  const depositReq  = Number(est['Deposit Required'] || est['depositRequired'] || 0);
  const isBWI       = !company || company.toLowerCase().includes('black wolf');
  const noDeposit   = depositType === 'none';

  if (noDeposit) {
    return 'Please review this estimate and use the buttons below to accept or decline. Full payment is due upon completion of the work.<br><br><a href="https://drive.google.com/file/d/1qo3vSEeG0_SlBxa-ahXALb5LnDNsyEEE/view?usp=drive_link" target="_blank">View Terms &amp; Conditions →</a>';
  }

  const depLabel = depositType === 'custom' && depositReq > 0
    ? `$${Math.round(depositReq).toLocaleString()} deposit`
    : '50% deposit';

  const cashNote = isBWI
    ? ' If paying cash or check we will make arrangements to pick up the deposit.'
    : '';

  return `Please review this estimate and use the buttons below to accept or decline. Once accepted, a payment link will be sent to collect the ${depLabel}.${cashNote} Balance is due upon job completion.<br><br><a href="https://drive.google.com/file/d/1qo3vSEeG0_SlBxa-ahXALb5LnDNsyEEE/view?usp=drive_link" target="_blank">View Terms &amp; Conditions →</a>`;
}

function buildProposalShellPdf(company,custName,custFirst,jobAddr,addrLine,estId,saved,expires,terms,contentHtml,photoUrls,companyAddr,companyPhone) {
  photoUrls = photoUrls||[];
  companyAddr  = companyAddr  || '169 Commack Rd #1009, Commack, NY 11725';
  companyPhone = companyPhone || '631-275-2020';
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>${estId}-Estimate</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Helvetica Neue',Arial,sans-serif;background:#fff;color:#222;font-size:14px;line-height:1.5}
  .page{max-width:780px;margin:0 auto;background:#fff;padding:52px 60px}
  .header{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:20px;margin-bottom:28px;border-bottom:3px solid #1B2B4B}
  .co-name{font-size:22px;font-weight:900;color:#1B2B4B;letter-spacing:-.4px;margin-bottom:4px}
  .co-sub{font-size:11px;color:#999;line-height:1.7}
  .bill-label{font-size:9px;font-weight:700;color:#999;text-transform:uppercase;letter-spacing:.12em;margin-bottom:5px;margin-top:14px}
  .cust-name{font-size:14px;font-weight:700;color:#1B2B4B}
  .cust-addr{font-size:12px;color:#555;line-height:1.5}
  .est-badge{font-size:34px;font-weight:900;color:#F26522;letter-spacing:-1px;line-height:1;margin-bottom:10px;text-align:right}
  .meta-table td{font-size:11px;padding:2px 0;vertical-align:top}
  .meta-table td:first-child{color:#1B2B4B;font-weight:700;padding-right:14px;text-align:right;white-space:nowrap}
  .meta-table td:last-child{color:#444;min-width:100px}
  .section-label{font-size:9px;font-weight:700;color:#1B2B4B;text-transform:uppercase;letter-spacing:.12em;margin-bottom:8px}
  .scope-box{background:#F7F8FA;border-left:3px solid #1B2B4B;border-radius:4px;padding:14px 18px;margin-bottom:20px;white-space:pre-wrap;font-size:12px;color:#333;line-height:1.7}
  table.line{width:100%;border-collapse:collapse;margin-bottom:0}
  table.line thead tr{background:#1B2B4B;color:#fff}
  table.line th{padding:11px 16px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em}
  table.line th:last-child{text-align:right}
  table.line td{padding:14px 16px;font-size:13px;border-bottom:2px solid #E2E5EA;vertical-align:top}
  table.line td:last-child{text-align:right;font-weight:800;color:#1B2B4B;font-size:14px;white-space:nowrap}
  .totals{display:flex;justify-content:flex-end;margin:16px 0 28px}
  .totals table{border-collapse:collapse;width:300px}
  .totals td{padding:5px 0;font-size:13px}
  .totals td:first-child{text-align:right;padding-right:16px;color:#555}
  .totals td:last-child{text-align:right;font-weight:700}
  .total-cash{font-size:20px;font-weight:900;color:#16A34A}
  .total-cc{font-size:13px;font-weight:600;color:#F26522;opacity:.85}
  .terms-box{background:#F7F8FA;border-left:3px solid #1B2B4B;border-radius:4px;padding:14px 18px;margin-bottom:24px}
  .terms-box p{font-size:12px;color:#444;line-height:1.8;white-space:pre-wrap}
  .opt-tabs{display:none}
  .opt-panel{display:block!important}
  .footer{display:flex;justify-content:space-between;align-items:center;padding-top:16px;border-top:1px solid #E2E5EA;font-size:11px;color:#aaa}
  .photo-grid img{width:100%;height:200px;object-fit:cover;border-radius:8px;border:1px solid #E2E5EA;display:block}
  /* hide all interactive/web-only sections in PDF view */
  .proceed-box,.action-section,.payment-step,.price-cards,.btn-accept,.btn-decline,
  .message-box,.submit-btn,.price-required-note,.response-done,
  .resp-accepted,.resp-declined,.decline-box,.how-to-proceed{display:none!important}
  /* print toolbar */
  #ptb{position:fixed;top:0;left:0;right:0;background:#1B2B4B;padding:10px 20px;display:flex;align-items:center;gap:12px;z-index:999;box-shadow:0 2px 8px rgba(0,0,0,.3)}
  #ptb span{font-size:12px;color:rgba(255,255,255,.6)}
  #ptb button{margin-left:auto;padding:8px 22px;border-radius:7px;border:none;cursor:pointer;font-size:13px;font-weight:800;background:#F26522;color:#fff}
  body{padding-top:52px}
  @media print{#ptb{display:none!important}body{padding-top:0!important}.page{padding:40px 50px!important;box-shadow:none!important}@page{margin:0.5in}}
</style>
</head>
<body>
<div id="ptb"><span>&#128196; &nbsp;Proposal — save or print using the button &rarr;</span><button onclick="window.print()">&#8595; Save as PDF</button></div>
<div class="page">
  <div class="header">
    <div>
      <div class="co-name">${escHtml(company)}</div>
      <div class="co-sub">${escHtml(companyAddr)}<br>${escHtml(companyPhone)}</div>
      <div class="bill-label">Bill To</div>
      <div class="cust-name">${escHtml(custName)}</div>
      <div class="cust-addr">${escHtml(jobAddr)}${addrLine?'<br>'+escHtml(addrLine):''}</div>
    </div>
    <div style="text-align:right;flex-shrink:0;margin-left:40px">
      <div class="est-badge">ESTIMATE</div>
      <table class="meta-table" style="margin-left:auto">
        <tr><td>Job #</td><td>${escHtml(estId)}</td></tr>
        <tr><td>Date</td><td>${saved}</td></tr>
        <tr><td>Expires</td><td>${expires}</td></tr>
      </table>
    </div>
  </div>
  ${contentHtml}
  ${photoUrls.length>0?`
  <div style="padding:0 0 24px">
    <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.08em;margin-bottom:12px">
      IMAGES &nbsp;<span style="font-weight:400;color:#aaa">For reference only — actual installation may vary slightly</span>
    </div>
    <div class="photo-grid" style="display:grid;grid-template-columns:${photoUrls.length===1?'1fr':photoUrls.length===2?'1fr 1fr':'repeat(2,1fr)'};gap:10px;">
      ${photoUrls.map(url=>{
        const src=url.includes('lh3.googleusercontent.com')?url:(()=>{const m=url.match(/[-\w]{25,}/);return m?'https://lh3.googleusercontent.com/d/'+m[0]:url;})();
        return '<img src="'+src+'" onerror="this.style.display=\'none\'">'; 
      }).join('')}
    </div>
  </div>`:''}
  <div class="footer">
    <span>© ${new Date().getFullYear()} ${escHtml(company)}</span>
    <span>This estimate is valid until ${expires}</span>
  </div>
</div>
</body>
</html>`;
}

// v2.16.0 - true when the estimate is set to Credit Card or a 0% financing term.
// Those estimates lead with the CC / Finance price; Cash / Check estimates are unchanged.
function isCardPay_(est) {
  return String((est && (est['Payment Method']||est['payMethod'])) || 'CASH').trim().toUpperCase() !== 'CASH';
}
// v2.16.0 - the price rows shown under the line items, lead price first.
function priceRowsHtml_(isCC, finalCC, finalCash, ccLabel, cashLabel) {
  const cc   = '$'+Math.round(finalCC).toLocaleString();
  const cash = '$'+Math.round(finalCash).toLocaleString();
  if (isCC) {
    return `<tr style="border-top:2px solid #1B2B4B"><td style="padding-top:12px;font-weight:800;color:#F26522">${ccLabel}</td><td style="padding-top:12px;font-size:20px;font-weight:900;color:#F26522">${cc}</td></tr>`
         + `<tr><td style="font-weight:600;color:#999;font-size:12px">${cashLabel}</td><td style="font-size:13px;font-weight:600;color:#16A34A;opacity:.85">${cash}</td></tr>`;
  }
  return `<tr style="border-top:2px solid #1B2B4B"><td style="padding-top:12px;font-weight:800;color:#16A34A">${cashLabel}</td><td style="padding-top:12px" class="total-cash">${cash}</td></tr>`;
}

function buildSingleEstimateContent(est, finalCC, finalCash, photoUrls) {
  const fenceType  = String(est['Fence Type']||'Fence Installation');
  const scopeNotes = String(est['Scope Notes']||'');
  const isCC       = isCardPay_(est);
  const leadAmt    = isCC ? finalCC : finalCash;
  const depositSentence = buildDepositTermsSentence(est);
  const proceedText     = buildProceedText(est);
  const terms = `${depositSentence}\n\nCustomer Obligations:\n1. Provide a clear area in the driveway for material to be delivered.\n2. Customer must fill out and sign NYS Capital Improvement form.\n3. We call 811 for all fence jobs. No exceptions.\n4. Customer is responsible for clearing fence line of any obstructions prior to fence installation.\n5. Any images attached to the proposal are for reference only and may be slightly different in person.`;
  return `
  ${scopeNotes?`<div class="section-label">Scope of Work</div><div class="scope-box">${escHtml(scopeNotes)}</div>`:''}
  <table class="line"><thead><tr><th style="width:30%">Item</th><th>Description</th><th style="width:130px">Amount</th></tr></thead>
  <tbody><tr><td style="font-weight:700;color:#1B2B4B">${escHtml(fenceType)}</td><td style="color:#555;font-size:12px">See scope of work above.</td><td>$${Math.round(leadAmt).toLocaleString()}</td></tr></tbody></table>
  <div class="totals"><table>
    ${priceRowsHtml_(isCC, finalCC, finalCash, 'CC / Finance Price', 'Cash / Check Price')}
  </table></div>
  <div class="section-label">Terms &amp; Obligations</div>
  <div class="terms-box"><p>${escHtml(terms)}</p></div>
  <div class="proceed-box"><div class="section-label" style="color:#F26522">How to Proceed</div>
  <p>${proceedText}</p></div>`;
}

function buildPhotoGrid(photoUrls) {
  if (!photoUrls||photoUrls.length===0) return '';
  const cols = photoUrls.length===1?'1fr':photoUrls.length===2?'1fr 1fr':'repeat(3,1fr)';
  const imgs = photoUrls.map(url=>{
    let src=url;
    if (!url.includes('lh3.googleusercontent.com')){const m=url.match(/[-\w]{25,}/);src=m?'https://lh3.googleusercontent.com/d/'+m[0]:url;}
    return `<img src="${escHtml(src)}" style="width:100%;height:auto;border-radius:6px;border:1px solid #E2E5EA;display:block;object-fit:cover" onerror="this.style.display='none'">`;
  }).join('');
  return `<div style="margin-bottom:20px"><div class="section-label">Site Images</div><div style="display:grid;grid-template-columns:${cols};gap:10px;margin-top:8px">${imgs}</div></div>`;
}

// v2.11.1 — job site per option. Read straight off an Estimates row.
function estSiteLineFromRow(est) {
  const a1   = String(est['Job Address']||'').trim();
  const a2   = String(est['Job Address 2']||'').trim();
  const city = String(est['Job City']||'').trim();
  const st   = String(est['Job State']||'').trim() || 'NY';
  const zip  = String(est['Zip']||'').trim().substring(0,5);
  if (!a1) return '';
  const tail = [city, [st, zip].filter(Boolean).join(' ')].filter(Boolean).join(', ');
  return [a1, a2, tail].filter(Boolean).join(', ');
}
// True only when two or more options have different, non-blank addresses.
function estSitesDifferRows(estimates) {
  const keys = {};
  (estimates||[]).forEach(e => { const l = estSiteLineFromRow(e); if (l) keys[l.toLowerCase().replace(/\s+/g,' ')] = 1; });
  return Object.keys(keys).length > 1;
}
// Returns '' unless the options are at different sites, so single-address
// proposals render exactly as they always have.
function estSiteHtmlRow(est, estimates, color, margin) {
  if (!estSitesDifferRows(estimates)) return '';
  const line = estSiteLineFromRow(est);
  if (!line) return '';
  return `<div style="font-size:12px;color:#333;margin-bottom:${margin||14}px"><span style="font-size:9px;font-weight:700;color:${color};text-transform:uppercase;letter-spacing:.12em">Job Site</span><br><strong>${escHtml(line)}</strong></div>`;
}

function buildOptionsContent(estimates) {
  const colors=['#F26522','#1B6FBF','#16A34A','#6D28D9','#0891B2'];
  // Use first estimate's deposit settings for terms (all options share same deposit)
  const depositSentence = buildDepositTermsSentence(estimates[0]||{});
  const proceedText     = buildProceedText(estimates[0]||{});
  const terms=`${depositSentence}\n\nCustomer Obligations:\n1. Provide a clear area in the driveway for material to be delivered.\n2. Customer must fill out and sign NYS Capital Improvement form.\n3. We call 811 for all fence jobs. No exceptions.\n4. Customer is responsible for clearing fence line of any obstructions prior to fence installation.\n5. Any images attached to the proposal are for reference only and may be slightly different in person.`;
  const tabs=estimates.map((est,i)=>{const letter=String(est['Estimate ID']||'').split('-')[2]||String.fromCharCode(65+i);const color=colors[i]||'#555';return `<button class="opt-tab${i===0?' active':''}" onclick="switchOptTab(${i})" style="${i===0?'color:'+color+';border-bottom-color:'+color:''}">Option ${letter}</button>`;}).join('');
  const panels=estimates.map((est,i)=>{const letter=String(est['Estimate ID']||'').split('-')[2]||String.fromCharCode(65+i);const color=colors[i]||'#555';const fenceType=String(est['Fence Type']||'Fence Installation');const scopeNotes=String(est['Scope Notes']||'');const finalCC=Number(est['Final Price CC'])||0;const finalCash=Number(est['Cash Price'])||0;
  return `<div class="opt-panel${i===0?' active':''}" id="opt_panel_${i}">${estSiteHtmlRow(est,estimates,color)}${scopeNotes?`<div class="section-label" style="color:${color}">Scope of Work — Option ${letter}</div><div class="scope-box" style="border-left-color:${color}">${escHtml(scopeNotes)}</div>`:''}${buildPhotoGrid(String(est['Photo URLs']||'').split(',').map(u=>u.trim()).filter(Boolean))}<table class="line"><thead><tr><th style="width:30%">Item</th><th>Description</th><th style="width:130px">Amount</th></tr></thead><tbody><tr><td style="font-weight:700;color:#1B2B4B">${escHtml(fenceType)}</td><td style="color:#555;font-size:12px">See scope of work above.</td><td>$${Math.round(isCardPay_(est)?finalCC:finalCash).toLocaleString()}</td></tr></tbody></table><div class="totals"><table>${priceRowsHtml_(isCardPay_(est), finalCC, finalCash, 'CC / Finance Price', 'Cash / Check Price')}</table></div></div>`;}).join('');
  return `<div class="opt-tabs">${tabs}</div>${panels}<div class="section-label">Terms &amp; Obligations</div><div class="terms-box"><p>${escHtml(terms)}</p></div><div class="proceed-box"><div class="section-label" style="color:#F26522">How to Proceed</div><p>Please review each option and select Accept to proceed with your chosen option. ${proceedText.replace("this estimate","each option")}</p></div>`;
}

function buildCombinedContent(estimates) {
  const colors=['#F26522','#1B6FBF','#16A34A','#6D28D9','#0891B2'];
  const depositSentence2 = buildDepositTermsSentence(estimates[0]||{});
  const proceedText2      = buildProceedText(estimates[0]||{});
  const terms=`${depositSentence2}\n\nCustomer Obligations:\n1. Provide a clear area in the driveway for material to be delivered.\n2. Customer must fill out and sign NYS Capital Improvement form.\n3. We call 811 for all fence jobs. No exceptions.\n4. Customer is responsible for clearing fence line of any obstructions prior to fence installation.\n5. Any images attached to the proposal are for reference only and may be slightly different in person.`;
  const combCC=Math.ceil(estimates.reduce((s,e)=>s+(Number(e['Final Price CC'])||0),0)/25)*25;
  const combCash=Math.ceil(estimates.reduce((s,e)=>s+(Number(e['Cash Price'])||0),0)/25)*25;
  const allCard=estimates.length>0 && estimates.every(isCardPay_);   // v2.16.0
  const scopeSections=estimates.map((est,i)=>{const letter=String(est['Estimate ID']||'').split('-')[2]||String.fromCharCode(65+i);const color=colors[i]||'#555';const fenceType=String(est['Fence Type']||'');const scopeNotes=String(est['Scope Notes']||'');const photoUrls=String(est['Photo URLs']||'').split(',').map(u=>u.trim()).filter(Boolean);return `<div class="scope-box" style="border-left-color:${color};margin-bottom:12px"><div style="font-size:9px;font-weight:700;color:${color};text-transform:uppercase;letter-spacing:.12em;margin-bottom:6px">Option ${letter} — ${escHtml(fenceType)}</div>${estSiteHtmlRow(est,estimates,color,8)}<div>${escHtml(scopeNotes)||'See scope details.'}</div></div>${buildPhotoGrid(photoUrls)}`;}).join('');
  const lineRows=estimates.map((est,i)=>{const letter=String(est['Estimate ID']||'').split('-')[2]||String.fromCharCode(65+i);const color=colors[i]||'#555';const fenceType=String(est['Fence Type']||'Fence Installation');const finalCash=allCard?(Number(est['Final Price CC'])||0):(Number(est['Cash Price'])||0);return `<tr><td style="vertical-align:top"><span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;color:${color};background:${color}18;margin-bottom:4px">Option ${letter}</span><div style="font-weight:700;color:#1B2B4B">${escHtml(fenceType)}</div></td><td style="color:#555;font-size:12px">See scope of work above.</td><td>$${Math.round(finalCash).toLocaleString()}</td></tr>`;}).join('');
  const subtotals=estimates.map((est,i)=>{const letter=String(est['Estimate ID']||'').split('-')[2]||String.fromCharCode(65+i);const finalCash=allCard?(Number(est['Final Price CC'])||0):(Number(est['Cash Price'])||0);return `<tr><td style="text-align:right;padding-right:16px;color:#999;font-size:11px">Option ${letter}</td><td style="text-align:right;font-size:11px;color:#666;font-weight:600">$${Math.round(finalCash).toLocaleString()}</td></tr>`;}).join('');
  return `<div class="section-label">Scope of Work</div>${scopeSections}<table class="line"><thead><tr><th style="width:30%">Item</th><th>Description</th><th style="width:130px">Amount</th></tr></thead><tbody>${lineRows}</tbody></table><div class="totals"><table>${subtotals}${allCard?priceRowsHtml_(true, combCC, combCash, 'Combined CC / Finance', 'Combined Cash / Check'):`<tr style="border-top:2px solid #1B2B4B"><td style="padding-top:12px;font-weight:800;color:#16A34A">Combined Cash / Check</td><td style="padding-top:12px" class="total-cash">$${Math.round(combCash).toLocaleString()}</td></tr>${estimates.some(isCardPay_)?`<tr><td style="font-weight:600;color:#999;font-size:12px">Combined CC / Finance</td><td class="total-cc" style="font-size:12px;color:#F26522;opacity:.85">$${Math.round(combCC).toLocaleString()}</td></tr>`:''}`}</table></div><div class="section-label">Terms &amp; Obligations</div><div class="terms-box"><p>${escHtml(terms)}</p></div><div class="proceed-box"><div class="section-label" style="color:#F26522">How to Proceed</div><p>Please review this combined estimate and use the buttons below to accept or decline. ${proceedText2}</p></div>`;
}

// ══════════════════════════════════════════════════════════════
// JOB COSTS
// ══════════════════════════════════════════════════════════════
function saveJobCost(data) {
  try {
    const sheet   = getOrCreateSheet(JOB_COSTS_TAB, JOB_COST_HEADERS);
    const costId  = String(data.costId||'').trim();
    const now     = new Date().toISOString();
    if (!costId) return { error:'No costId provided' };
    const row = [costId, String(data.jobId||'').trim(), String(data.estimateId||'').trim(), data.type||'', data.desc||'', data.date||'', Number(data.amount)||0, data.createdAt||now, now, String(data.vendor||'').trim()];
    const existingRow = findRow(sheet, costId, 0);
    if (existingRow > 0) {
      const existing = sheet.getRange(existingRow,1,1,JOB_COST_HEADERS.length).getValues()[0];
      row[7] = existing[7]||now;
      sheet.getRange(existingRow,1,1,row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { success:true, costId };
  } catch(e) { return { error:e.toString() }; }
}

function deleteJobCost(data) {
  try {
    const sheet  = getOrCreateSheet(JOB_COSTS_TAB, JOB_COST_HEADERS);
    const costId = String(data.costId||'').trim();
    if (!costId) return { error:'No costId provided' };
    const values = sheet.getDataRange().getValues();
    for (let i = values.length-1; i >= 1; i--) {
      if (String(values[i][0]).trim()===costId) { sheet.deleteRow(i+1); return { success:true, deleted:costId }; }
    }
    return { error:'Cost not found: '+costId };
  } catch(e) { return { error:e.toString() }; }
}

function getJobCosts() {
  try {
    const sheet   = getOrCreateSheet(JOB_COSTS_TAB, JOB_COST_HEADERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { costs:[] };
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const costs = data.slice(1).filter(row=>row[0]).map(row=>{
      const obj={};headers.forEach((h,i)=>{let v=row[i];if(v instanceof Date)v=isNaN(v.getTime())?'':v.toISOString().split('T')[0];obj[h]=v;});
      return {costId:String(obj['Cost ID']||''),jobId:String(obj['Job ID']||''),estimateId:String(obj['Estimate ID']||''),type:String(obj['Type']||''),desc:String(obj['Description']||''),date:String(obj['Date']||''),amount:Number(obj['Amount'])||0,createdAt:String(obj['Created At']||''),updatedAt:String(obj['Updated At']||''),vendor:String(obj['Vendor']||'')};
    });
    return { costs, count:costs.length };
  } catch(e) { return { costs:[], error:e.toString() }; }
}

// ══════════════════════════════════════════════════════════════
// ANALYTICS
// ══════════════════════════════════════════════════════════════
const ANALYTICS_TAB = 'Analytics Data';

function saveAnalytics(data) {
  try {
    const records = data.records;
    if (!records||!records.length) return { error:'No records provided' };
    const sheet = getOrCreateSheet(ANALYTICS_TAB, []);
    sheet.clearContents();
    const headers = Object.keys(records[0]);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    const rows = records.map(r=>headers.map(h=>r[h]||''));
    sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
    sheet.getRange(1,headers.length+1).setValue('_uploaded_at');
    sheet.getRange(2,headers.length+1).setValue(new Date().toISOString());
    return { success:true, count:records.length };
  } catch(e) { return { error:e.toString() }; }
}

function getAnalytics() {
  try {
    const sheet = getSpreadsheet().getSheetByName(ANALYTICS_TAB);
    if (!sheet) return { records:[], empty:true };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { records:[], empty:true };
    const allHeaders = data[0].map(h=>String(h).trim());
    const metaIdx    = allHeaders.indexOf('_uploaded_at');
    const headers    = metaIdx>=0?allHeaders.slice(0,metaIdx):allHeaders;
    const tz = Session.getScriptTimeZone();
    // Sheets auto-converts date-like text (e.g. "8/15/2026") into a real
    // Date-typed cell on write (same as typing it in by hand) — that's
    // normal Sheets behavior. The bug was reading it back out: a naive
    // String(cell) on a Date object produces JS's verbose default format
    // ("Sat Aug 15 2026 00:00:00 GMT-0400 (Eastern Daylight Time)"), which
    // the front end's parseAnalyticsDate() can't recognize as ISO or
    // M/D/YYYY. Every date-range filter except "All Time" then silently
    // dropped every row on any device other than the one that originally
    // uploaded the CSV (which never round-trips through this — it keeps
    // its own in-memory copy of the original CSV text). Fix: format any
    // real Date cell explicitly as ISO (yyyy-MM-dd) before sending it out,
    // which the front end already fully supports.
    const records = data.slice(1).filter(r=>r.slice(0,headers.length).some(c=>String(c).trim())).map(r=>{
      const obj={};
      headers.forEach((h,i)=>{
        const cell = r[i];
        obj[h] = (cell instanceof Date) ? Utilities.formatDate(cell, tz, 'yyyy-MM-dd') : String(cell||'').trim();
      });
      return obj;
    });
    const uploadedAt = metaIdx>=0&&data[1]?String(data[1][metaIdx]||''):'';
    return { records, count:records.length, uploadedAt };
  } catch(e) { return { error:e.toString() }; }
}

// ══════════════════════════════════════════════════════════════
// SETUP DATA
// ══════════════════════════════════════════════════════════════
function saveSetupData(data) {
  if (!data||!data.type) return { error:'Missing type' };
  const sheet   = getOrCreateSheet(SETUP_DATA_TAB, ['Type','Payload','Updated At']);
  const values  = sheet.getDataRange().getValues();
  const now     = new Date().toISOString();
  const payload = typeof data.payload==='string'?data.payload:JSON.stringify(data.payload);
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]===data.type) { sheet.getRange(i+1,2,1,2).setValues([[payload,now]]); return { success:true, action:'updated', type:data.type, savedAt:now }; }
  }
  sheet.appendRow([data.type, payload, now]);
  return { success:true, action:'created', type:data.type, savedAt:now };
}

function getSetupData() {
  try {
    const sheet  = getOrCreateSheet(SETUP_DATA_TAB, ['Type','Payload','Updated At']);
    const values = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < values.length; i++) {
      const type=values[i][0],payload=values[i][1];
      if (!type) continue;
      try { result[type]=JSON.parse(payload); } catch(e) { result[type]=payload; }
    }
    return { setupData:result };
  } catch(e) { return { setupData:{} }; }
}

// v2.8.0 — resolves a company code (BWI/TRF/BATPIG/...) to its real
// name/address/phone from the same Setup ▸ Manage Companies data the
// front end uses (this same Setup_Data sheet, type 'companies'). Used
// anywhere a Google-Apps-Script-rendered page (the hosted proposal page,
// its PDF, the owner notification email) needs company info server-side,
// where there's no front-end CI_COMPANIES lookup to reach for.
// Falls back to BWI's known values if the company list hasn't been
// saved yet, or the code doesn't match any saved company (e.g. a token
// generated before that company existed in Setup).
function getCompanyRecord(code) {
  const FALLBACK = {
    name: 'Black Wolf Industries LLC',
    addr1: '169 Commack Rd #1009', addr2: '', city: 'Commack', state: 'NY', zip: '11725',
    phone: '631-275-2020'
  };
  try {
    const setup = getSetupData().setupData || {};
    const list  = Array.isArray(setup.companies) ? setup.companies : [];
    const rec   = list.find(c => c && c.code === (code || 'BWI'))
               || list.find(c => c && c.code === 'BWI')
               || FALLBACK;
    const addr  = [rec.addr1, rec.addr2, [rec.city, rec.state, rec.zip].filter(Boolean).join(', ')]
                    .filter(Boolean).join(', ')
               || [FALLBACK.addr1, FALLBACK.addr2, [FALLBACK.city, FALLBACK.state, FALLBACK.zip].filter(Boolean).join(', ')]
                    .filter(Boolean).join(', ');
    return {
      name:  rec.name  || FALLBACK.name,
      addr:  addr,
      phone: rec.phone || FALLBACK.phone
    };
  } catch (e) {
    return {
      name:  FALLBACK.name,
      addr:  [FALLBACK.addr1, [FALLBACK.city, FALLBACK.state, FALLBACK.zip].filter(Boolean).join(', ')].filter(Boolean).join(', '),
      phone: FALLBACK.phone
    };
  }
}

// ══════════════════════════════════════════════════════════════
// ZIP LOOKUP — self-growing zip → city/county table
// Replaces the old hardcoded NASSAU_ZIPS/SUFFOLK_ZIPS JS objects.
// Seeded once with the original 152 Nassau/Suffolk entries; grows
// whenever an unknown zip is typed in on the Estimator or Contact
// form and that form is saved (see saveZipEntry, called from
// saveJob/saveContact below).
// ══════════════════════════════════════════════════════════════
function getZipLookup() {
  try {
    const sheet  = getOrCreateSheet(ZIP_LOOKUP_TAB, ZIP_LOOKUP_HEADERS);
    const values = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < values.length; i++) {
      const zip = String(values[i][0]).trim();
      if (!zip) continue;
      result[zip] = { city: values[i][1] || '', county: values[i][2] || '' };
    }
    return { zipLookup: result };
  } catch(e) { return { zipLookup: {} }; }
}

// Dedup write — only inserts if the zip isn't already known.
// Called from saveJob() and saveContact(), not on every keystroke,
// so abandoned forms never pollute the sheet.
function saveZipEntry(zip, city, county) {
  zip = String(zip || '').trim();
  city = String(city || '').trim();
  county = String(county || '').trim();
  if (!zip || zip.length !== 5 || (!city && !county)) return;

  try {
    const sheet  = getOrCreateSheet(ZIP_LOOKUP_TAB, ZIP_LOOKUP_HEADERS);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < existing.length; i++) {
        if (String(existing[i][0]).trim() === zip) return; // already known — skip
      }
    }
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setNumberFormat('@'); // force Zip column to text — preserves leading zeros
    sheet.appendRow([zip, city, county, new Date().toISOString()]);
  } catch(e) {
    // Non-fatal — zip lookup write failures should never block a job/contact save
  }
}

// One-time seed — populates Zip_Lookup with the original 156
// Nassau/Suffolk entries from the old hardcoded JS tables, so the
// sheet starts pre-filled instead of empty. Safe to run multiple
// times — skips zips that already exist. Run manually from the
// Apps Script editor once after deploying this version; not
// called automatically from the app.
function seedZipLookupFromLegacyTables() {
  const sheet = getOrCreateSheet(ZIP_LOOKUP_TAB, ZIP_LOOKUP_HEADERS);
  const lastRow = sheet.getLastRow();
  const existingZips = {};
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existing.forEach(r => { if (r[0]) existingZips[String(r[0]).trim()] = true; });
  }

  const legacyData = [
    ['11001','Floral Park','Nassau'],
    ['11002','Floral Park','Nassau'],
    ['11003','Elmont','Nassau'],
    ['11004','Glen Oaks','Nassau'],
    ['11005','Floral Park','Nassau'],
    ['11010','Franklin Square','Nassau'],
    ['11020','Great Neck','Nassau'],
    ['11021','Great Neck','Nassau'],
    ['11022','Great Neck','Nassau'],
    ['11023','Great Neck','Nassau'],
    ['11024','Great Neck','Nassau'],
    ['11030','Manhasset','Nassau'],
    ['11040','New Hyde Park','Nassau'],
    ['11042','New Hyde Park','Nassau'],
    ['11050','Port Washington','Nassau'],
    ['11096','Inwood','Nassau'],
    ['11101','Long Island City','Nassau'],
    ['11102','Astoria','Nassau'],
    ['11103','Astoria','Nassau'],
    ['11104','Sunnyside','Nassau'],
    ['11105','Astoria','Nassau'],
    ['11106','Astoria','Nassau'],
    ['11109','Long Island City','Nassau'],
    ['11701','Amityville','Suffolk'],
    ['11702','Babylon','Suffolk'],
    ['11703','North Babylon','Suffolk'],
    ['11704','West Babylon','Suffolk'],
    ['11705','Bayport','Suffolk'],
    ['11706','Bay Shore','Suffolk'],
    ['11707','West Babylon','Suffolk'],
    ['11709','Bayville','Suffolk'],
    ['11710','Bellmore','Suffolk'],
    ['11713','Bellport','Suffolk'],
    ['11714','Bethpage','Suffolk'],
    ['11715','Blue Point','Suffolk'],
    ['11716','Bohemia','Suffolk'],
    ['11717','Brentwood','Suffolk'],
    ['11718','Brightwaters','Suffolk'],
    ['11719','Brookhaven','Suffolk'],
    ['11720','Centereach','Suffolk'],
    ['11721','Centerport','Suffolk'],
    ['11722','Central Islip','Suffolk'],
    ['11724','Cold Spring Harbor','Suffolk'],
    ['11725','Commack','Suffolk'],
    ['11726','Copiague','Suffolk'],
    ['11727','Coram','Suffolk'],
    ['11729','Deer Park','Suffolk'],
    ['11730','East Islip','Suffolk'],
    ['11731','East Northport','Suffolk'],
    ['11732','East Norwich','Suffolk'],
    ['11733','East Setauket','Suffolk'],
    ['11735','Farmingdale','Suffolk'],
    ['11737','Farmingdale','Suffolk'],
    ['11738','Farmingville','Suffolk'],
    ['11739','Great River','Suffolk'],
    ['11740','Greenlawn','Suffolk'],
    ['11741','Holbrook','Suffolk'],
    ['11742','Holtsville','Suffolk'],
    ['11743','Huntington','Suffolk'],
    ['11746','Huntington Station','Suffolk'],
    ['11747','Melville','Suffolk'],
    ['11749','Islandia','Suffolk'],
    ['11751','Islip','Suffolk'],
    ['11752','Islip Terrace','Suffolk'],
    ['11753','Jericho','Suffolk'],
    ['11754','Kings Park','Suffolk'],
    ['11755','Lake Grove','Suffolk'],
    ['11756','Levittown','Suffolk'],
    ['11757','Lindenhurst','Suffolk'],
    ['11758','Massapequa','Suffolk'],
    ['11760','Melville','Suffolk'],
    ['11762','Massapequa Park','Suffolk'],
    ['11763','Medford','Suffolk'],
    ['11764','Miller Place','Suffolk'],
    ['11765','Mill Neck','Suffolk'],
    ['11766','Mount Sinai','Suffolk'],
    ['11767','Nesconset','Suffolk'],
    ['11768','Northport','Suffolk'],
    ['11769','Oakdale','Suffolk'],
    ['11770','Ocean Beach','Suffolk'],
    ['11771','Oyster Bay','Suffolk'],
    ['11772','Patchogue','Suffolk'],
    ['11773','Syosset','Suffolk'],
    ['11775','Melville','Suffolk'],
    ['11776','Port Jefferson Station','Suffolk'],
    ['11777','Port Jefferson','Suffolk'],
    ['11778','Rocky Point','Suffolk'],
    ['11779','Ronkonkoma','Suffolk'],
    ['11780','Saint James','Suffolk'],
    ['11782','Sayville','Suffolk'],
    ['11783','Seaford','Suffolk'],
    ['11784','Selden','Suffolk'],
    ['11786','Shoreham','Suffolk'],
    ['11787','Smithtown','Suffolk'],
    ['11788','Hauppauge','Suffolk'],
    ['11789','Sound Beach','Suffolk'],
    ['11790','Stony Brook','Suffolk'],
    ['11791','Syosset','Suffolk'],
    ['11792','Wading River','Suffolk'],
    ['11793','Wantagh','Suffolk'],
    ['11794','Stony Brook','Suffolk'],
    ['11795','West Islip','Suffolk'],
    ['11796','West Sayville','Suffolk'],
    ['11797','Woodbury','Suffolk'],
    ['11798','Wyandanch','Suffolk'],
    ['11801','Hicksville','Suffolk'],
    ['11802','Hicksville','Suffolk'],
    ['11803','Plainview','Suffolk'],
    ['11804','Old Bethpage','Suffolk'],
    ['11815','Hicksville','Suffolk'],
    ['11853','Jericho','Suffolk'],
    ['11901','Riverhead','Suffolk'],
    ['11930','Amagansett','Suffolk'],
    ['11931','Aquebogue','Suffolk'],
    ['11932','Bridgehampton','Suffolk'],
    ['11933','Calverton','Suffolk'],
    ['11934','Center Moriches','Suffolk'],
    ['11935','Cutchogue','Suffolk'],
    ['11937','East Hampton','Suffolk'],
    ['11939','East Marion','Suffolk'],
    ['11940','East Moriches','Suffolk'],
    ['11941','Eastport','Suffolk'],
    ['11942','East Quogue','Suffolk'],
    ['11944','Greenport','Suffolk'],
    ['11946','Hampton Bays','Suffolk'],
    ['11947','Jamesport','Suffolk'],
    ['11948','Laurel','Suffolk'],
    ['11949','Manorville','Suffolk'],
    ['11950','Mastic','Suffolk'],
    ['11951','Mastic Beach','Suffolk'],
    ['11952','Mattituck','Suffolk'],
    ['11953','Middle Island','Suffolk'],
    ['11954','Montauk','Suffolk'],
    ['11955','Moriches','Suffolk'],
    ['11956','New Suffolk','Suffolk'],
    ['11957','Orient','Suffolk'],
    ['11958','Peconic','Suffolk'],
    ['11959','Quogue','Suffolk'],
    ['11960','Remsenburg','Suffolk'],
    ['11961','Ridge','Suffolk'],
    ['11962','Sagaponack','Suffolk'],
    ['11963','Sag Harbor','Suffolk'],
    ['11964','Shelter Island','Suffolk'],
    ['11965','Shelter Island Heights','Suffolk'],
    ['11967','Shirley','Suffolk'],
    ['11968','Southampton','Suffolk'],
    ['11969','Southampton','Suffolk'],
    ['11970','South Jamesport','Suffolk'],
    ['11971','Southold','Suffolk'],
    ['11972','Speonk','Suffolk'],
    ['11973','Upton','Suffolk'],
    ['11975','Wainscott','Suffolk'],
    ['11976','Water Mill','Suffolk'],
    ['11977','Westhampton','Suffolk'],
    ['11978','Westhampton Beach','Suffolk'],
    ['11980','Yaphank','Suffolk']
  ];

  const now = new Date().toISOString();
  const rowsToAppend = [];
  legacyData.forEach(([zip, city, county]) => {
    if (!existingZips[zip]) rowsToAppend.push([zip, city, county, now]);
  });

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, 1).setNumberFormat('@'); // force Zip column to text
    sheet.getRange(startRow, 1, rowsToAppend.length, 4).setValues(rowsToAppend);
  }
  Logger.log('Seeded ' + rowsToAppend.length + ' zip entries (' + (legacyData.length - rowsToAppend.length) + ' already existed).');
  return { seeded: rowsToAppend.length, skipped: legacyData.length - rowsToAppend.length };
}


// ══════════════════════════════════════════════════════════════
// GET JOB FILES
// ══════════════════════════════════════════════════════════════
function getJobFiles(data) {
  if (!data||!data.jobId) return { error:'Missing jobId' };
  try {
    const folder = getJobFolder(data.jobId, data.custName||'');
    if (!folder) return { files:[], folderUrl:null, empty:true };
    const folderUrl = folder.getUrl();
    const result = [];
    const fmtSize = bytes => bytes<1024?bytes+' B':bytes<1048576?(bytes/1024).toFixed(0)+' KB':(bytes/1048576).toFixed(1)+' MB';
    const mimeIcon = mime => {if(!mime)return '??';if(mime.includes('pdf'))return '??';if(mime.includes('image'))return '??️';if(mime.includes('spreadsheet')||mime.includes('excel'))return '??';if(mime.includes('document')||mime.includes('word'))return '??';if(mime.includes('video'))return '??';return '??';};
    const rootFiles = folder.getFiles();
    while (rootFiles.hasNext()) { const f=rootFiles.next(); result.push({name:f.getName(),url:f.getUrl(),mime:f.getMimeType(),icon:mimeIcon(f.getMimeType()),size:fmtSize(f.getSize()),updatedAt:f.getLastUpdated().toISOString(),subfolder:'(root)'}); }
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) { const sub=subfolders.next(); const subName=sub.getName(); const subFiles=sub.getFiles(); while(subFiles.hasNext()){const f=subFiles.next();result.push({name:f.getName(),url:f.getUrl(),mime:f.getMimeType(),icon:mimeIcon(f.getMimeType()),size:fmtSize(f.getSize()),updatedAt:f.getLastUpdated().toISOString(),subfolder:subName});} }
    result.sort((a,b)=>a.subfolder<b.subfolder?-1:a.subfolder>b.subfolder?1:a.name.localeCompare(b.name));
    return { files:result, folderUrl, count:result.length };
  } catch(e) { return { error:e.message, files:[], folderUrl:null }; }
}

// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════
function errorPage(msg) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Proposal — Not Available</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Helvetica Neue',Arial,sans-serif;background:#F7F8FA;display:flex;align-items:center;justify-content:center;min-height:100vh;color:#444}.box{text-align:center;padding:48px;max-width:420px}.icon{font-size:48px;margin-bottom:16px}.h{font-size:20px;font-weight:800;color:#1B2B4B;margin-bottom:10px}.p{font-size:13px;color:#666;line-height:1.6}.contact{margin-top:20px;font-size:13px}.contact a{color:#F26522;font-weight:700;text-decoration:none}</style></head><body><div class="box"><div class="icon">??</div><div class="h">Proposal Not Available</div><div class="p">${msg}</div><div class="contact">Questions? Call us at <a href="tel:6312752020">631-275-2020</a></div></div></body></html>`;
}

function escHtml(str) {
  return String(str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function escJs(str) {
  return String(str||'').replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'\\"');
}

function fmt(n) {
  return '$' + Math.round(Number(n)||0).toLocaleString();
}

function getImageAsBase64(data) {
  try {
    let fileId = data.fileId||'';
    if (!fileId&&data.url) { const m=data.url.match(/[-\w]{25,}/); if(m) fileId=m[0]; }
    if (!fileId) return { error:'No file ID found' };
    const file=DriveApp.getFileById(fileId);
    const blob=file.getBlob();
    const mimeType=blob.getContentType()||'image/jpeg';
    const b64=Utilities.base64Encode(blob.getBytes());
    return { success:true, dataUrl:'data:'+mimeType+';base64,'+b64 };
  } catch(e) { return { error:e.toString() }; }
}

// ══════════════════════════════════════════════════════════════
// GET ALL — single startup call that replaces 5 separate requests
// Returns jobs, invoices, jobCosts, settings, and setupData in one round trip
// ══════════════════════════════════════════════════════════════
function getAll() {
  try {
    const jobsResult     = getJobs();
    const invoicesResult = getInvoices();
    const costsResult    = getJobCosts();
    const settingsResult = getSettings();
    const setupResult    = getSetupData();
    const zipResult      = getZipLookup();
    return {
      jobs:      jobsResult.jobs      || [],
      invoices:  invoicesResult.invoices || [],
      jobCosts:  costsResult.costs    || [],
      settings:  settingsResult.settings || {},
      setupData: setupResult.setupData || {},
      zipLookup: zipResult.zipLookup   || {},
      counts: {
        jobs:     jobsResult.count     || 0,
        invoices: invoicesResult.count || 0,
        costs:    costsResult.count    || 0
      }
    };
  } catch(e) {
    return { error: e.toString(), jobs: [], invoices: [], jobCosts: [], settings: {}, setupData: {}, zipLookup: {} };
  }
}

// ══════════════════════════════════════════════════════════════
// GOOGLE CALENDAR INTEGRATION
// ══════════════════════════════════════════════════════════════
// v2.9.1 — Didz asked for the phone number and address in a calendar
// event's description to be actual clickable links (tap to dial, tap to
// open in Maps/Waze) instead of plain text. The Calendar API's description
// field officially supports HTML, so this builds real <a href="tel:...">
// and <a href="https://www.google.com/maps/...">  tags. Shared by both
// calendar-writing functions below (createCalendarEvent for job installs,
// createEstimateAppointment for estimate visits) so the behavior — and any
// future fix to it — stays in exactly one place.
function _calDescEscape(s) {
  // Minimal HTML-escaping for text that gets embedded alongside real <a>
  // tags in an HTML description — a customer name or note containing "&"
  // or "<" should never be mistaken for markup.
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function _calPhoneLink(phone) {
  const digits = String(phone || '').replace(/\D/g, '');
  if (!digits) return '';
  // Best-effort US number handling — 10 digits gets a +1, an 11-digit
  // number already starting with 1 just gets the +. Anything else (already
  // has a country code, or is malformed) is passed through as-is rather
  // than guessed at.
  const e164 = digits.length === 10 ? '1' + digits : digits;
  return '<a href="tel:+' + e164 + '">📞 ' + _calDescEscape(phone) + '</a>';
}
function _calMapsLink(addr) {
  const clean = String(addr || '').trim();
  if (!clean) return '';
  const url = 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(clean);
  return '<a href="' + url + '">📍 ' + _calDescEscape(clean) + '</a>';
}

// Creates a calendar event on the owner's primary Google Calendar
// when a job is scheduled. Called from index.html after updateStatus
// succeeds. Returns the event ID so the client can store it.
//
// data fields expected:
//   jobId          — e.g. "937"
//   estimateId     — e.g. "937-Frank-C-2026"
//   custName       — e.g. "Frank C"
//   custPhone      — e.g. "(631) 555-0142" — v2.9.1, becomes a tap-to-dial link
//   fenceType      — e.g. "6ft Wood Privacy"
//   jobAddr        — e.g. "123 Main St"
//   jobCity        — e.g. "Smithtown"
//   jobState       — e.g. "NY"
//   scheduledDate  — "YYYY-MM-DD"
//   scheduledTime  — "Morning" | "Afternoon" | "Full Day"
//   scheduledCrew  — e.g. "Crew A"
//   scheduleNotes  — free text
function createCalendarEvent(data) {
  try {
    // ── Build title: [937] Frank C - 6ft Wood Privacy ──────
    const jobNum    = String(data.jobId || '').trim();
    const custName  = String(data.custName  || '').trim();
    const fenceType = String(data.fenceType || '').trim();
    const title     = '[' + jobNum + '] ' + custName + (fenceType ? ' - ' + fenceType : '');

    // ── Build location ─────────────────────────────────────
    const addr      = [
      String(data.jobAddr  || '').trim(),
      String(data.jobCity  || '').trim(),
      String(data.jobState || '').trim()
    ].filter(Boolean).join(', ');

    // ── Determine start/end times from slot ───────────────
    const dateStr = String(data.scheduledDate || '').trim(); // "YYYY-MM-DD"
    if (!dateStr) return { error: 'No scheduledDate provided' };

    const slot = String(data.scheduledTime || 'Full Day').trim();
    const dateParts = dateStr.split('-').map(Number); // [YYYY, MM, DD]
    if (dateParts.length < 3) return { error: 'Invalid scheduledDate format' };

    let startHour, endHour;
    if (slot === 'Morning') {
      startHour = 8;  endHour = 12;
    } else if (slot === 'Afternoon') {
      startHour = 13; endHour = 17;
    } else {
      // Full Day (default)
      startHour = 8;  endHour = 17;
    }

    // Apps Script month is 0-indexed
    const startTime = new Date(dateParts[0], dateParts[1] - 1, dateParts[2], startHour, 0, 0);
    const endTime   = new Date(dateParts[0], dateParts[1] - 1, dateParts[2], endHour,   0, 0);

    // ── Build description ──────────────────────────────────
    const crew  = String(data.scheduledCrew  || '').trim();
    const notes = String(data.scheduleNotes  || '').trim();
    const estId = String(data.estimateId     || '').trim();
    const phoneLink = _calPhoneLink(data.custPhone);
    const mapsLink  = _calMapsLink(addr);
    let desc = 'Job: ' + _calDescEscape(estId || jobNum);
    if (phoneLink) desc += '\n' + phoneLink;
    if (mapsLink)  desc += '\n' + mapsLink;
    if (crew)  desc += '\nCrew: ' + _calDescEscape(crew);
    if (notes) desc += '\nNotes: ' + _calDescEscape(notes);
    desc += '\n\nCreated by FencePro';

    // ── Create the event on the primary calendar ───────────
    const cal = CalendarApp.getDefaultCalendar();

    // v2.10.0 — rescheduling a job used to always create a brand-new event
    // and leave the old one sitting on the calendar (calendarEventId was
    // never stored anywhere, so there was nothing to delete). Same fix
    // pattern already used by createEstimateAppointment().
    const oldEventId = String(data.oldEventId || '').trim();
    if (oldEventId) {
      try {
        const oldEvent = cal.getEventById(oldEventId);
        if (oldEvent) oldEvent.deleteEvent();
      } catch(e) {
        Logger.log('createCalendarEvent: could not delete old event ' + oldEventId + ': ' + e.toString());
      }
    }

    const event = cal.createEvent(title, startTime, endTime, {
      location:    addr,
      description: desc
    });

    const eventId  = event.getId();
    const eventUrl = 'https://calendar.google.com/calendar/r/eventedit/' +
                     encodeURIComponent(eventId);

    Logger.log('Calendar event created: ' + eventId + ' — ' + title);
    return { success: true, eventId, eventUrl, title, slot };

  } catch(e) {
    Logger.log('createCalendarEvent error: ' + e.toString());
    return { error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// ESTIMATE APPOINTMENT SCHEDULING (v2.9.0)
// ══════════════════════════════════════════════════════════════
// Books an initial estimate/site-visit appointment for a Contact, before
// any job/estimate record exists yet. Deliberately a separate function
// from createCalendarEvent() above (which is wired to the job-install
// scheduling flow's date/slot/crew shape and already works in production)
// so this new feature can't regress that one. Both write to the same
// CalendarApp.getDefaultCalendar(), so everything still lands on the one
// calendar Didz already checks.
//
// data fields expected:
//   custName    — contact's name
//   phone       — contact's phone — v2.9.1, becomes a tap-to-dial link in the description
//   addr, city, state, zip — becomes the event location (and, v2.9.1, a tap-to-open-in-Maps link)
//   dateTime    — "YYYY-MM-DDTHH:mm" (specific start time, not a date-only slot)
//   notes       — free text
//   oldEventId  — optional; if set (rescheduling), the previous event is
//                 deleted first so it doesn't leave a duplicate/stale entry
//                 on the calendar. Safe to omit or to pass an ID that no
//                 longer exists — deletion failures are swallowed.
function createEstimateAppointment(data) {
  try {
    const custName = String(data.custName || '').trim() || 'Unknown';
    const dateTimeStr = String(data.dateTime || '').trim(); // "YYYY-MM-DDTHH:mm"
    if (!dateTimeStr) return { error: 'No dateTime provided' };

    const [datePart, timePart] = dateTimeStr.split('T');
    const dateParts = (datePart || '').split('-').map(Number); // [YYYY, MM, DD]
    const timeParts = (timePart || '10:00').split(':').map(Number); // [HH, mm]
    if (dateParts.length < 3) return { error: 'Invalid dateTime format' };

    // Apps Script month is 0-indexed. Default appointment length: 1 hour.
    const startTime = new Date(dateParts[0], dateParts[1] - 1, dateParts[2], timeParts[0] || 0, timeParts[1] || 0, 0);
    const endTime   = new Date(startTime.getTime() + 60 * 60 * 1000);

    const addr = [
      String(data.addr  || '').trim(),
      String(data.city  || '').trim(),
      String(data.state || '').trim(),
      String(data.zip   || '').trim(),
    ].filter(Boolean).join(', ');

    const phone = String(data.phone || '').trim();
    const notes = String(data.notes || '').trim();
    const phoneLink = _calPhoneLink(phone);
    const mapsLink  = _calMapsLink(addr);
    let desc = 'Estimate visit for ' + _calDescEscape(custName);
    if (phoneLink) desc += '\n' + phoneLink;
    if (mapsLink)  desc += '\n' + mapsLink;
    if (notes) desc += '\nNotes: ' + _calDescEscape(notes);
    desc += '\n\nCreated by FencePro';

    const cal = CalendarApp.getDefaultCalendar();

    // Rescheduling — remove the old event first so it doesn't linger.
    const oldEventId = String(data.oldEventId || '').trim();
    if (oldEventId) {
      try {
        const oldEvent = cal.getEventById(oldEventId);
        if (oldEvent) oldEvent.deleteEvent();
      } catch(e) {
        Logger.log('createEstimateAppointment: could not delete old event ' + oldEventId + ': ' + e.toString());
      }
    }

    const title = 'Estimate: ' + custName;
    const event = cal.createEvent(title, startTime, endTime, {
      location:    addr,
      description: desc
    });

    const eventId  = event.getId();
    const eventUrl = 'https://calendar.google.com/calendar/r/eventedit/' +
                     encodeURIComponent(eventId);

    Logger.log('Estimate appointment created: ' + eventId + ' — ' + title);
    return { success: true, eventId, eventUrl, title };

  } catch(e) {
    Logger.log('createEstimateAppointment error: ' + e.toString());
    return { error: e.toString() };
  }
}

// Cancels a previously-booked estimate appointment by calendar event ID.
// Idempotent — if the event is already gone (manually deleted, etc.) this
// still reports success, since the end state (no calendar event) is the
// same either way.
function cancelEstimateAppointment(data) {
  try {
    const eventId = String((data && data.eventId) || '').trim();
    if (!eventId) return { error: 'Missing eventId' };
    const cal   = CalendarApp.getDefaultCalendar();
    const event = cal.getEventById(eventId);
    if (event) event.deleteEvent();
    return { success: true };
  } catch(e) {
    Logger.log('cancelEstimateAppointment error: ' + e.toString());
    return { error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// MIGRATION — adds Scheduled Time column to live Jobs sheet
// Run ONCE from the Apps Script editor (not via the web app):
//   1. Open Apps Script editor
//   2. Select migrateAddScheduledTime from the function dropdown
//   3. Click Run
// Safe to run multiple times — skips if column already exists.
// ══════════════════════════════════════════════════════════════
function migrateAddScheduledTime() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_TAB);
  if (!sheet) {
    Logger.log('ERROR: Jobs tab not found');
    SpreadsheetApp.getUi().alert('Migration Error', 'Jobs tab not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existing  = headerRow.map(h => String(h).trim());

  if (existing.includes('Scheduled Time')) {
    Logger.log('Scheduled Time column already exists — no action needed.');
    SpreadsheetApp.getUi().alert('Migration', 'Scheduled Time column already exists — nothing to do.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Insert after Scheduled Crew
  const crewIdx = existing.indexOf('Scheduled Crew');
  if (crewIdx === -1) {
    Logger.log('ERROR: Scheduled Crew column not found');
    SpreadsheetApp.getUi().alert('Migration Error', 'Could not find Scheduled Crew column.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // insertColumnAfter is 1-based
  sheet.insertColumnAfter(crewIdx + 1);
  const newCol = crewIdx + 2; // column index after insertion
  sheet.getRange(1, newCol).setValue('Scheduled Time')
    .setBackground('#1a1a1a').setFontColor('#e8c547')
    .setFontWeight('bold').setFontSize(10).setFontFamily('Courier New');

  sheet.setFrozenRows(1);
  Logger.log('Added Scheduled Time column at position ' + newCol);
  SpreadsheetApp.getUi().alert(
    'Migration Complete',
    'Scheduled Time column added to Jobs tab at column ' + newCol + '.\n\nExisting data was not modified.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ══════════════════════════════════════════════════════════════
// MIGRATION — adds Job City, Job State, Job Zip, Customer Note
// columns to the live Invoices sheet.
// Run ONCE from the Apps Script editor:
//   1. Select migrateAddInvoiceFields from the function dropdown
//   2. Click Run
// Safe to run multiple times — skips columns that already exist.
// ══════════════════════════════════════════════════════════════
function migrateAddInvoiceFields() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVOICES_TAB);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Migration Error', 'Invoices tab not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existing  = headerRow.map(h => String(h).trim());

  // Columns to add: { header, insertAfter } — insertAfter is the column it goes after
  const toAdd = [
    { header: 'Job City',      insertAfter: 'Job Address'    },
    { header: 'Job State',     insertAfter: 'Job City'       },
    { header: 'Job Zip',       insertAfter: 'Job State'      },
    { header: 'Customer Note', insertAfter: 'Payment Instructions' },
  ];

  let added = 0;
  // Re-read headers after each insertion so indexes stay accurate
  toAdd.forEach(({ header, insertAfter }) => {
    const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (current.includes(header)) {
      Logger.log(header + ' already exists — skipping');
      return;
    }
    const afterIdx = current.indexOf(insertAfter);
    if (afterIdx === -1) {
      Logger.log('WARNING: could not find anchor column "' + insertAfter + '" for "' + header + '" — appending at end');
      const endCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, endCol).setValue(header)
        .setBackground('#1B2B4B').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(10);
    } else {
      sheet.insertColumnAfter(afterIdx + 1); // 1-based
      const newCol = afterIdx + 2;
      sheet.getRange(1, newCol).setValue(header)
        .setBackground('#1B2B4B').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(10);
      Logger.log('Added "' + header + '" at column ' + newCol);
    }
    added++;
  });

  sheet.setFrozenRows(1);
  const msg = added > 0
    ? added + ' column(s) added to Invoices tab. Existing data was not modified.'
    : 'All columns already present — nothing to do.';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Invoice Migration Complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function keepWarm() {
  try { getSpreadsheet().getName(); console.log('keepWarm: OK at '+new Date().toISOString()); }
  catch(e) { console.warn('keepWarm: failed — '+e.message); }
}

// ══════════════════════════════════════════════════════════════
// MIGRATIONS (run once from Apps Script editor)
// ══════════════════════════════════════════════════════════════
// ── RUN THIS FIRST after deploying v2.3 ──────────────────────
// Fixes the Jobs tab header row in production — adds any missing
// columns without touching existing data or reordering anything.
function migrateJobsSchema() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_TAB);
  if (!sheet) { Logger.log('ERROR: Jobs tab not found'); return { error: 'Jobs tab missing' }; }
  const existingHdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>String(h).trim());
  Logger.log('Current Jobs headers (' + existingHdrs.length + '): ' + existingHdrs.join(' | '));
  let added = 0;
  JOB_HEADERS.forEach(h => {
    if (!existingHdrs.includes(h)) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(h)
        .setBackground('#1a1a1a').setFontColor('#e8c547')
        .setFontWeight('bold').setFontSize(10).setFontFamily('Courier New');
      existingHdrs.push(h);
      added++;
      Logger.log('Added missing column: ' + h + ' at col ' + nextCol);
    }
  });
  sheet.setFrozenRows(1);
  const msg = added > 0
    ? 'Added ' + added + ' missing column(s) to Jobs tab.'
    : 'Jobs tab schema already up to date.';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Jobs Schema Migration', msg + '\n\nTotal columns: ' + sheet.getLastColumn(), SpreadsheetApp.getUi().ButtonSet.OK);
  return { success: true, addedColumns: added, totalColumns: sheet.getLastColumn() };
}

function migrateJobIds() {
  const estSheet = getOrCreateSheet(ESTIMATES_TAB, ESTIMATE_HEADERS);
  const jobSheet = getOrCreateSheet(JOBS_TAB, JOB_HEADERS);
  const estData    = estSheet.getDataRange().getValues();
  const estHeaders = estData[0].map(h=>String(h).trim());
  const jobIdCol   = estHeaders.indexOf('Job ID');
  const estIdCol   = estHeaders.indexOf('Estimate ID');
  const jobNoCol   = estHeaders.indexOf('Job No');
  if (jobIdCol===-1) { Logger.log('ERROR: Job ID column not found'); return { error:'Job ID column missing' }; }
  let estRowsUpdated=0;
  for (let i=1;i<estData.length;i++) {
    const estimateId=String(estData[i][estIdCol]||'').trim();
    const currentJobId=String(estData[i][jobIdCol]||'').trim();
    if (!estimateId) continue;
    const extractedJobId=estimateId.split('-')[0];
    if (!extractedJobId) continue;
    if (!currentJobId) { estSheet.getRange(i+1,jobIdCol+1).setValue(extractedJobId); estData[i][jobIdCol]=extractedJobId; estRowsUpdated++; }
    if (jobNoCol!==-1) { const currentJobNo=String(estData[i][jobNoCol]||'').trim(); if(currentJobNo===estimateId||currentJobNo===''){estSheet.getRange(i+1,jobNoCol+1).setValue(extractedJobId);estData[i][jobNoCol]=extractedJobId;} }
  }
  const jobData    = jobSheet.getDataRange().getValues();
  const jobHeaders = jobData.length>1?jobData[0].map(h=>String(h).trim()):JOB_HEADERS;
  const existingJobMap={};
  if (jobData.length>1) { jobData.slice(1).forEach(row=>{const key=String(row[0]||'').trim();if(!key)return;const jobId=key.split('-')[0];if(!existingJobMap[jobId]){const obj={};jobHeaders.forEach((h,k)=>{obj[h]=row[k];});existingJobMap[jobId]=obj;}});}
  const uniqueJobIds={};
  estData.slice(1).forEach(row=>{const estimateId=String(row[estIdCol]||'').trim();const jobId=String(row[jobIdCol]||estimateId.split('-')[0]||'').trim();if(!jobId)return;if(!uniqueJobIds[jobId]){const estObj={};estHeaders.forEach((h,k)=>{estObj[h]=row[k];});uniqueJobIds[jobId]={contactUID:String(estObj['Contact UID']||''),custName:String(estObj['Customer Name']||'')};}});
  const lastJobRow=jobSheet.getLastRow();
  if (lastJobRow>1) jobSheet.getRange(2,1,lastJobRow-1,JOB_HEADERS.length).clearContent();
  let jobsWritten=0;
  Object.entries(uniqueJobIds).forEach(([jobId,seed])=>{const existing=existingJobMap[jobId]||{};const newRow=[jobId,existing['Estimate ID']||'',existing['Contact UID']||seed.contactUID||'',existing['Customer Name']||seed.custName||'',existing['Status']||'Draft',existing['Status Changed At']||new Date().toISOString(),existing['Changed By']||'Migration',Number(existing['Contract Price'])||0,existing['Contract Price Type']||'',existing['Deposit Type']||'50pct',Number(existing['Deposit Required'])||0,Number(existing['Deposit Paid'])||0,existing['Deposit Method']||'',existing['Deposit Date']||'',existing['Deposit Ref']||'',existing['Scheduled Date']||'',existing['Scheduled Crew']||'',existing['Schedule Notes']||'',Number(existing['Final Payment Amount'])||0,existing['Final Payment Method']||'',existing['Final Payment Date']||'',existing['Final Payment Ref']||'',Number(existing['Balance Due'])||0,existing['Lost Reason']||'',existing['Notes']||''];jobSheet.appendRow(newRow);jobsWritten++;});
  const summary={success:true,estimatesBackfilled:estRowsUpdated,jobsWritten,message:`Migration complete. ${estRowsUpdated} estimate rows backfilled. ${jobsWritten} job rows written.`};
  Logger.log(JSON.stringify(summary));
  return summary;
}

function migrateEstimatesSchema() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Estimates');
  if (!sheet) { Logger.log('ERROR: Estimates tab not found'); return; }
  const headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const existing  = headerRow.map(h=>String(h).trim());
  const EXPECTED  = ['Estimate ID','Job Prefix','Job No','Contact UID','Saved At','Customer Name','Phone','Email','Job Address','Job Address 2','Job City','Job State','Zip','County','Job Type','Fence Type','Payment Method','Supplier','Total LF','Num Gates','Markup Pct','Franchise On','Material Cost','Labor Cost','Demo LF','Demo Rate','Concrete Bags','Carting Fee','AddOn Costs','Discount','Supplier CC Rate','Delivery Fee','Mat Del PreTax','Material Tax','Supplier Fee','Concrete Total','Base SM Price','Effective Markup Pct','Final Price CC','Cash Price','Customer Tax','Final With Tax','Gross Margin Pct','Gross Margin Dollars','Material Pct','Labor Pct','Scope Notes','Project Description','Delivery Override','Channel','Campaign','Photos Count','Drive Folder','Internal Notes','Photo URLs','Doc URLs','Job ID','Status','Deposit Type','Deposit Required','Contract Price','Contract Price Type','Fence-Only Price Per LF','Material'];
  let added=0;
  EXPECTED.forEach((header,i)=>{const col=i+1;if(existing[i]!==header){if(col>sheet.getLastColumn()){sheet.getRange(1,col).setValue(header);}else if(!existing.includes(header)){sheet.insertColumnBefore(col);sheet.getRange(1,col).setValue(header);}added++;}});
  const lastCol=Math.max(sheet.getLastColumn(),EXPECTED.length);
  const headerRange=sheet.getRange(1,1,1,lastCol);
  headerRange.setFontWeight('bold').setBackground('#1B2B4B').setFontColor('#FFFFFF').setWrap(false);
  Logger.log('Migration complete. '+added+' columns added/fixed.');
  SpreadsheetApp.getUi().alert('Migration Complete',added+' column(s) added to Estimates tab.\n\nTotal columns: '+sheet.getLastColumn()+'\n\nYour existing data was not modified.',SpreadsheetApp.getUi().ButtonSet.OK);
}

function resetCards() {
  // Utility — clears cached card state if needed
  Logger.log('resetCards called');
}


// ============================================================
// VENDOR PARTS CATALOG + BILLS OF MATERIALS  (v2.17.0)
// ------------------------------------------------------------
// Vendor_Parts: one row per part. Every part gets its own unique Part ID
// (VP-0001, VP-0002, ...) that never changes and is what everything else
// (bills of materials, later the job link) points at. The Part Number is a
// separate, editable field, and it must be unique across the whole catalog
// (case-insensitive) so the same part number can never be entered twice.
// Deleting a part only marks it inactive so old bills of materials keep
// their data.
//
// Bills_Of_Materials: one row per bill. Lines are stored as a JSON snapshot
// (part, price and quantities as they were when saved). The Job ID column is
// intentionally left blank for now; it is where the job link will go later.
// ============================================================
const VENDOR_PARTS_TAB = 'Vendor_Parts';
const VP_FIELDS = [
  ['partId','Part ID'], ['partNumber','Part Number'], ['vendor','Vendor'], ['price','Price'],
  ['type','Type'], ['subType','Sub Type'], ['description','Description'], ['color','Color'],
  ['fenceHeight','Fence Height'], ['height','Height'], ['width','Width'],
  ['unitStd','Unit Std'], ['unitDivisor','Unit Divisor'], ['roundingQty','Rounding Qty'],
  ['active','Active'], ['createdAt','Created At'], ['updatedAt','Updated At']
];
const VENDOR_PART_HEADERS = VP_FIELDS.map(function(f){ return f[1]; });

const BOM_TAB = 'Bills_Of_Materials';
const BOM_HEADERS = ['BOM ID','Label','Job ID','Status','Notes','Total','Lines JSON','Created At','Updated At','Estimate ID','Customer Name'];

let _vpFormatted = false;

function _vpSheet_() {
  const sheet = getOrCreateSheet(VENDOR_PARTS_TAB, VENDOR_PART_HEADERS);
  if (!_vpFormatted) {
    // Part numbers like 12409 or 1E5 must stay text, never turn into numbers or dates.
    const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
    ['Part ID', 'Part Number'].forEach(function(name) {
      const c = hdr.indexOf(name);
      if (c >= 0) sheet.getRange(2, c + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
    });
    _vpFormatted = true;
  }
  return sheet;
}

function _vpKey_(partNumber) { return String(partNumber == null ? '' : partNumber).trim().toLowerCase(); }

function _vpNumOrBlank_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(String(v).replace(/[$,\s]/g, ''));
  return isNaN(n) ? NaN : n;
}

// Read the whole catalog once. Returns header info plus parsed rows.
function _vpReadAll_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  const idx = {};
  hdr.forEach(function(h, i){ idx[h] = i; });
  const rows = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    values.forEach(function(r, i) {
      if (!String(r[idx['Part ID']] || '').trim() && !String(r[idx['Part Number']] || '').trim()) return;
      rows.push({ rowNum: i + 2, obj: _vpRowToObj_(idx, r) });
    });
  }
  return { hdr: hdr, idx: idx, rows: rows };
}

function _vpDateStr_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
  return String(v || '');
}

function _vpRowToObj_(idx, r) {
  function g(name) { return idx[name] === undefined ? '' : r[idx[name]]; }
  const num = function(v) { const n = _vpNumOrBlank_(v); return (typeof n === 'number' && !isNaN(n)) ? n : ''; };
  return {
    partId:      String(g('Part ID')).trim(),
    partNumber:  String(g('Part Number')).trim(),
    vendor:      String(g('Vendor')).trim(),
    price:       Number(g('Price')) || 0,
    type:        String(g('Type')).trim(),
    subType:     String(g('Sub Type')).trim(),
    description: String(g('Description')).trim(),
    color:       String(g('Color')).trim(),
    fenceHeight: num(g('Fence Height')),
    height:      num(g('Height')),
    width:       num(g('Width')),
    unitStd:     String(g('Unit Std')).trim() || 'EA',
    unitDivisor: Number(g('Unit Divisor')) > 0 ? Number(g('Unit Divisor')) : 1,
    roundingQty: Number(g('Rounding Qty')) > 0 ? Number(g('Rounding Qty')) : 1,
    active:      String(g('Active')).trim().toLowerCase() !== 'no',
    createdAt:   _vpDateStr_(g('Created At')),
    updatedAt:   _vpDateStr_(g('Updated At'))
  };
}

function _vpObjToRow_(hdr, obj, baseRow) {
  const row = baseRow ? baseRow.slice() : new Array(hdr.length).fill('');
  VP_FIELDS.forEach(function(f) {
    const c = hdr.indexOf(f[1]);
    if (c < 0) return;
    let v = obj[f[0]];
    if (f[0] === 'active') v = (v === false) ? 'No' : 'Yes';
    row[c] = (v === null || v === undefined) ? '' : v;
  });
  return row;
}

// Validate + clean one incoming part. Returns { clean } or { error }.
function _vpClean_(d) {
  d = d || {};
  const partNumber = String(d.partNumber == null ? '' : d.partNumber).trim();
  const vendor     = String(d.vendor == null ? '' : d.vendor).trim();
  if (!partNumber) return { error: 'Part number is required' };
  if (!vendor)     return { error: 'Vendor is required' };
  const priceRaw = (d.price === null || d.price === undefined) ? '' : String(d.price).replace(/[$,\s]/g, '');
  if (priceRaw === '') return { error: 'Price is required' };
  const price = Number(priceRaw);
  if (isNaN(price) || price < 0) return { error: 'Price must be a number, zero or higher' };

  const dims = {};
  const labels = { fenceHeight: 'Fence height', height: 'Height', width: 'Width' };
  for (const k in labels) {
    const n = _vpNumOrBlank_(d[k]);
    if (typeof n === 'number' && isNaN(n)) return { error: labels[k] + ' must be a number' };
    dims[k] = n;
  }
  let unitDivisor = d.unitDivisor === '' || d.unitDivisor == null ? 1 : Number(d.unitDivisor);
  let roundingQty = d.roundingQty === '' || d.roundingQty == null ? 1 : Number(d.roundingQty);
  if (isNaN(unitDivisor) || unitDivisor <= 0) return { error: 'Unit divisor must be greater than zero' };
  if (isNaN(roundingQty) || roundingQty <= 0) return { error: 'Rounding qty must be greater than zero' };

  const unitStd = String(d.unitStd == null ? '' : d.unitStd).trim().toUpperCase().replace(/\.+$/, '') || 'EA';
  return { clean: {
    partNumber: partNumber, vendor: vendor, price: Math.round(price * 10000) / 10000,
    type: String(d.type == null ? '' : d.type).trim().toUpperCase(),
    subType: String(d.subType == null ? '' : d.subType).trim().toUpperCase(),
    description: String(d.description == null ? '' : d.description).trim(),
    color: String(d.color == null ? '' : d.color).trim().toUpperCase(),
    fenceHeight: dims.fenceHeight, height: dims.height, width: dims.width,
    unitStd: unitStd, unitDivisor: unitDivisor, roundingQty: roundingQty
  } };
}

function _vpNextNum_(rows, prefix) {
  let max = 0;
  rows.forEach(function(r) {
    const m = String(r.obj ? r.obj.partId : r).match(/^VP-(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return max;
}

function _vpId_(n) { return 'VP-' + ('0000' + n).slice(-4); }

function getVendorParts() {
  try {
    const sheet = _vpSheet_();
    const all = _vpReadAll_(sheet);
    const parts = all.rows.map(function(r){ return r.obj; });
    return { parts: parts, count: parts.length };
  } catch (e) { return { parts: [], error: e.toString() }; }
}

// Add or edit one part. data.partId present = edit that part.
function saveVendorPart(data) {
  const c = _vpClean_(data);
  if (c.error) return { error: c.error };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const sheet = _vpSheet_();
    const all = _vpReadAll_(sheet);
    const key = _vpKey_(c.clean.partNumber);
    const editId = String(data.partId || '').trim();
    let target = null;
    if (editId) {
      target = all.rows.filter(function(r){ return r.obj.partId === editId; })[0];
      if (!target) return { error: 'Part not found: ' + editId };
    }
    const clash = all.rows.filter(function(r){
      return _vpKey_(r.obj.partNumber) === key && (!target || r.obj.partId !== target.obj.partId);
    })[0];
    if (clash) {
      return { error: 'Part number "' + clash.obj.partNumber + '" already exists (' + clash.obj.partId + ', vendor ' + clash.obj.vendor + ')' +
                      (clash.obj.active ? '' : '. It is inactive; restore it instead of adding it again.'),
               duplicate: true, existingPartId: clash.obj.partId };
    }
    const now = new Date().toISOString();
    let obj;
    if (target) {
      obj = Object.assign({}, target.obj, c.clean);
      if (data.active === true || data.active === false) obj.active = data.active;
      obj.updatedAt = now;
      const base = sheet.getRange(target.rowNum, 1, 1, all.hdr.length).getValues()[0];
      sheet.getRange(target.rowNum, 1, 1, all.hdr.length).setValues([_vpObjToRow_(all.hdr, obj, base)]);
    } else {
      obj = Object.assign({ partId: _vpId_(_vpNextNum_(all.rows) + 1), active: true, createdAt: now, updatedAt: now }, c.clean);
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, all.hdr.length).setValues([_vpObjToRow_(all.hdr, obj)]);
    }
    return { success: true, part: _vpRowToObj_(all.idx, _vpObjToRow_(all.hdr, obj)) };
  } catch (e) { return { error: e.toString() };
  } finally { lock.releaseLock(); }
}

// Import many parts at once. data.parts = array, data.onDuplicate = 'skip' | 'update'.
function bulkSaveVendorParts(data) {
  const incoming = (data && data.parts) || [];
  if (!incoming.length) return { error: 'No parts provided' };
  const mode = (data.onDuplicate === 'update') ? 'update' : 'skip';
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const sheet = _vpSheet_();
    const all = _vpReadAll_(sheet);
    const byKey = {};
    all.rows.forEach(function(r){ byKey[_vpKey_(r.obj.partNumber)] = r; });
    let nextNum = _vpNextNum_(all.rows);
    const now = new Date().toISOString();
    const newRows = [];
    const updates = [];
    const rejected = [];
    let skipped = 0;
    incoming.forEach(function(p, i) {
      const c = _vpClean_(p);
      if (c.error) { rejected.push({ row: i + 1, partNumber: String((p && p.partNumber) || ''), reason: c.error }); return; }
      const key = _vpKey_(c.clean.partNumber);
      const hit = byKey[key];
      if (hit) {
        if (mode === 'update' && !hit.pending) {
          const obj = Object.assign({}, hit.obj, c.clean, { updatedAt: now });
          hit.obj = obj;
          if (!hit.queued) { hit.queued = true; updates.push(hit); }
        } else { skipped++; }
        return;
      }
      nextNum++;
      const obj = Object.assign({ partId: _vpId_(nextNum), active: true, createdAt: now, updatedAt: now }, c.clean);
      const rec = { rowNum: 0, obj: obj, pending: true };
      byKey[key] = rec;
      newRows.push(_vpObjToRow_(all.hdr, obj));
    });
    updates.forEach(function(u) {
      const base = sheet.getRange(u.rowNum, 1, 1, all.hdr.length).getValues()[0];
      sheet.getRange(u.rowNum, 1, 1, all.hdr.length).setValues([_vpObjToRow_(all.hdr, u.obj, base)]);
    });
    if (newRows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, all.hdr.length).setValues(newRows);
    }
    return { success: true, added: newRows.length, updated: updates.length, skipped: skipped, rejected: rejected };
  } catch (e) { return { error: e.toString() };
  } finally { lock.releaseLock(); }
}

// Soft delete: the part stays in the sheet marked inactive.
function deleteVendorPart(data) {
  const partId = String((data && data.partId) || '').trim();
  if (!partId) return { error: 'No partId provided' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const sheet = _vpSheet_();
    const all = _vpReadAll_(sheet);
    const hit = all.rows.filter(function(r){ return r.obj.partId === partId; })[0];
    if (!hit) return { error: 'Part not found: ' + partId };
    hit.obj.active = false;
    hit.obj.updatedAt = new Date().toISOString();
    const base = sheet.getRange(hit.rowNum, 1, 1, all.hdr.length).getValues()[0];
    sheet.getRange(hit.rowNum, 1, 1, all.hdr.length).setValues([_vpObjToRow_(all.hdr, hit.obj, base)]);
    return { success: true, part: hit.obj };
  } catch (e) { return { error: e.toString() };
  } finally { lock.releaseLock(); }
}

// ── BILLS OF MATERIALS ─────────────────────────────────────
// Pack math, shared by the screen and the server so they can never disagree:
//   orderUnits = qty rounded up to a multiple of Rounding Qty
//   packs      = orderUnits / Unit Divisor, rounded up
//   lineTotal  = packs x price      (price is per pack of Unit Divisor units)
function _bomCalc_(line) {
  const qty = Number(line.qty) || 0;
  const div = Number(line.unitDivisor) > 0 ? Number(line.unitDivisor) : 1;
  const rnd = Number(line.roundingQty) > 0 ? Number(line.roundingQty) : 1;
  const price = Number(line.price) || 0;
  const orderUnits = qty > 0 ? Math.ceil(qty / rnd - 1e-9) * rnd : 0;
  const packs = orderUnits > 0 ? Math.ceil(orderUnits / div - 1e-9) : 0;
  return { orderUnits: orderUnits, packs: packs, lineTotal: Math.round(packs * price * 100) / 100 };
}

function _bomNextNum_(sheet) {
  const last = sheet.getLastRow();
  let max = 0;
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 1).getValues().forEach(function(r) {
      const m = String(r[0]).match(/^BOM-(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return max;
}

function _bomRowToObj_(hdr, r) {
  const g = function(n) { const c = hdr.indexOf(n); return c < 0 ? '' : r[c]; };
  let lines = [];
  try { lines = JSON.parse(String(g('Lines JSON')) || '[]'); } catch (e) { lines = []; }
  return {
    bomId: String(g('BOM ID')).trim(), label: String(g('Label')), jobId: String(g('Job ID')).trim(),
    status: String(g('Status')) || 'Draft', notes: String(g('Notes')), total: Number(g('Total')) || 0,
    lines: lines, createdAt: _vpDateStr_(g('Created At')), updatedAt: _vpDateStr_(g('Updated At')),
    estimateId: String(g('Estimate ID')).trim(), customerName: String(g('Customer Name'))
  };
}

function getBOMs() {
  try {
    const sheet = getOrCreateSheet(BOM_TAB, BOM_HEADERS);
    const last = sheet.getLastRow();
    if (last < 2) return { boms: [], count: 0 };
    const lastCol = sheet.getLastColumn();
    const hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
    const boms = sheet.getRange(2, 1, last - 1, lastCol).getValues()
      .filter(function(r){ return String(r[0]).trim(); })
      .map(function(r){ return _bomRowToObj_(hdr, r); });
    return { boms: boms, count: boms.length };
  } catch (e) { return { boms: [], error: e.toString() }; }
}

// Add or edit a bill. data.bomId present = edit that bill.
function saveBOM(data) {
  data = data || {};
  const lines = Array.isArray(data.lines) ? data.lines : [];
  if (!lines.length) return { error: 'Add at least one part before saving' };
  const clean = [];
  let total = 0;
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i] || {};
    if (!String(l.partId || '').trim() || !String(l.partNumber || '').trim()) return { error: 'Line ' + (i + 1) + ' has no part' };
    if (!(Number(l.qty) > 0)) return { error: 'Line ' + (i + 1) + ' (' + l.partNumber + ') needs a quantity greater than zero' };
    const calc = _bomCalc_(l);
    total += calc.lineTotal;
    clean.push({
      partId: String(l.partId).trim(), partNumber: String(l.partNumber).trim(), vendor: String(l.vendor || '').trim(),
      description: String(l.description || ''), color: String(l.color || ''), unitStd: String(l.unitStd || 'EA'),
      unitDivisor: Number(l.unitDivisor) > 0 ? Number(l.unitDivisor) : 1,
      roundingQty: Number(l.roundingQty) > 0 ? Number(l.roundingQty) : 1,
      price: Number(l.price) || 0, qty: Number(l.qty),
      orderUnits: calc.orderUnits, packs: calc.packs, lineTotal: calc.lineTotal
    });
  }
  total = Math.round(total * 100) / 100;
  const json = JSON.stringify(clean);
  if (json.length > 45000) return { error: 'This bill is too large to save (' + clean.length + ' lines). Split it into two bills.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const sheet = getOrCreateSheet(BOM_TAB, BOM_HEADERS);
    const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
    const now = new Date().toISOString();
    const editId = String(data.bomId || '').trim();
    let rowNum = -1, existing = null;
    if (editId) {
      rowNum = findRow(sheet, editId, hdr.indexOf('BOM ID'));
      if (rowNum < 0) return { error: 'Bill of materials not found: ' + editId };
      existing = sheet.getRange(rowNum, 1, 1, hdr.length).getValues()[0];
    }
    const bomId = editId || ('BOM-' + ('0000' + (_bomNextNum_(sheet) + 1)).slice(-4));
    const isTpl = String(data.status || '') === 'Template';   // templates are never linked to a job
    const vals = {
      'BOM ID': bomId, 'Label': String(data.label || '').trim() || bomId, 'Job ID': isTpl ? '' : String(data.jobId || '').trim(),
      'Estimate ID': isTpl ? '' : String(data.estimateId || '').trim(), 'Customer Name': isTpl ? '' : String(data.customerName || '').trim(),
      'Status': String(data.status || 'Draft'), 'Notes': String(data.notes || ''), 'Total': total, 'Lines JSON': json,
      'Created At': existing ? (existing[hdr.indexOf('Created At')] || now) : now, 'Updated At': now
    };
    const row = existing ? existing.slice() : new Array(hdr.length).fill('');
    BOM_HEADERS.forEach(function(h) { const c = hdr.indexOf(h); if (c >= 0) row[c] = vals[h]; });
    if (rowNum > 0) sheet.getRange(rowNum, 1, 1, hdr.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, hdr.length).setValues([row]);
    return { success: true, bom: _bomRowToObj_(hdr, row) };
  } catch (e) { return { error: e.toString() };
  } finally { lock.releaseLock(); }
}

function deleteBOM(data) {
  const bomId = String((data && data.bomId) || '').trim();
  if (!bomId) return { error: 'No bomId provided' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const sheet = getOrCreateSheet(BOM_TAB, BOM_HEADERS);
    const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
    const rowNum = findRow(sheet, bomId, hdr.indexOf('BOM ID'));
    if (rowNum < 0) return { error: 'Bill of materials not found: ' + bomId };
    sheet.deleteRow(rowNum);
    return { success: true, deleted: bomId };
  } catch (e) { return { error: e.toString() };
  } finally { lock.releaseLock(); }
}


// ============================================================
// SUPPLIER CONTACTS  (v2.19.0)
// One row per person. A supplier with several people has several rows,
// so one order email can go to all of them. Supplier names match the
// Vendor column of Vendor_Parts (compared case-insensitively).
// ============================================================
const SC_TAB = 'Supplier_Contacts';
const SC_HEADERS = ['Contact ID','Supplier','Contact Name','Email','Active','Created At','Updated At'];
function _scRead_() {
  const sheet = getOrCreateSheet(SC_TAB, SC_HEADERS);
  const last = sheet.getLastRow();
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  const out = [];
  if (last >= 2) {
    const vals = sheet.getRange(2, 1, last - 1, hdr.length).getValues();
    vals.forEach(function(r, i) {
      const g = function(n) { const c = hdr.indexOf(n); return c < 0 ? '' : r[c]; };
      if (!String(g('Contact ID')).trim()) return;
      const act = String(g('Active')).trim().toLowerCase();
      out.push({ rowNum: i + 2, obj: {
        contactId: String(g('Contact ID')).trim(), supplier: String(g('Supplier')).trim(),
        name: String(g('Contact Name')).trim(), email: String(g('Email')).trim(),
        active: !(act === 'no' || act === 'false'),
        createdAt: _vpDateStr_(g('Created At')), updatedAt: _vpDateStr_(g('Updated At'))
      }});
    });
  }
  return { sheet: sheet, hdr: hdr, rows: out };
}
function _scRowOf_(hdr, o, base) {
  const row = base ? base.slice() : new Array(hdr.length).fill('');
  const m = { 'Contact ID': o.contactId, 'Supplier': o.supplier, 'Contact Name': o.name, 'Email': o.email,
    'Active': o.active ? 'Yes' : 'No', 'Created At': o.createdAt, 'Updated At': o.updatedAt };
  Object.keys(m).forEach(function(h) { const c = hdr.indexOf(h); if (c >= 0) row[c] = m[h]; });
  return row;
}
function getSupplierContacts() {
  try { return { success: true, contacts: _scRead_().rows.map(function(r){ return r.obj; }) }; }
  catch (e) { return { error: e.toString() }; }
}
function saveSupplierContact(data) {
  data = data || {};
  const supplier = String(data.supplier || '').trim();
  const name = String(data.name || '').trim();
  const email = String(data.email || '').trim();
  if (!supplier) return { error: 'Supplier is required' };
  if (!name) return { error: 'Contact name is required' };
  if (!email) return { error: 'Email is required' };
  if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) return { error: 'That email address does not look right' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const all = _scRead_();
    const editId = String(data.contactId || '').trim();
    const dup = all.rows.filter(function(r) {
      return r.obj.active && r.obj.contactId !== editId &&
        r.obj.supplier.toLowerCase() === supplier.toLowerCase() && r.obj.email.toLowerCase() === email.toLowerCase();
    })[0];
    if (dup) return { error: dup.obj.name + ' already has that email under ' + dup.obj.supplier, duplicate: true };
    const now = new Date().toISOString();
    let obj, rowNum = -1, base = null;
    if (editId) {
      const hit = all.rows.filter(function(r){ return r.obj.contactId === editId; })[0];
      if (!hit) return { error: 'Contact not found: ' + editId };
      obj = hit.obj; rowNum = hit.rowNum;
      base = all.sheet.getRange(rowNum, 1, 1, all.hdr.length).getValues()[0];
    } else {
      let max = 0;
      all.rows.forEach(function(r) { const m = r.obj.contactId.match(/^SC-(\d+)$/); if (m) max = Math.max(max, parseInt(m[1], 10)); });
      obj = { contactId: 'SC-' + ('0000' + (max + 1)).slice(-4), createdAt: now, active: true };
    }
    obj.supplier = supplier; obj.name = name; obj.email = email; obj.updatedAt = now;
    const row = _scRowOf_(all.hdr, obj, base);
    if (rowNum > 0) all.sheet.getRange(rowNum, 1, 1, all.hdr.length).setValues([row]);
    else all.sheet.getRange(all.sheet.getLastRow() + 1, 1, 1, all.hdr.length).setValues([row]);
    return { success: true, contact: obj };
  } catch (e) { return { error: e.toString() };
  } finally { lock.releaseLock(); }
}
function deleteSupplierContact(data) {
  const id = String((data && data.contactId) || '').trim();
  if (!id) return { error: 'No contactId provided' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { error: 'The server is busy. Please try again.' }; }
  try {
    const all = _scRead_();
    const hit = all.rows.filter(function(r){ return r.obj.contactId === id; })[0];
    if (!hit) return { error: 'Contact not found: ' + id };
    hit.obj.active = false; hit.obj.updatedAt = new Date().toISOString();
    const base = all.sheet.getRange(hit.rowNum, 1, 1, all.hdr.length).getValues()[0];
    all.sheet.getRange(hit.rowNum, 1, 1, all.hdr.length).setValues([_scRowOf_(all.hdr, hit.obj, base)]);
    return { success: true, contact: hit.obj };
  } catch (e) { return { error: e.toString() };
  } finally { lock.releaseLock(); }
}
