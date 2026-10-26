# Hearth Café Ordering and Management System

**Revised target:** VB.NET Windows Forms (.NET 8) + Microsoft Access (`.accdb`) through `System.Data.OleDb`, aligned to the uploaded desktop-application rubric.

## Included modules

- First-run administrator account creation and login/logout
- PBKDF2-SHA256 password hashing with per-user random salts
- Role-based navigation and action checks (`ADMIN`, `STAFF`)
- Menu Management: add, edit, search, availability, archive/restore, required-field and duplicate validation
- New Order: menu selection, cart, quantities, discounts, payment validation, transactional order + item persistence, receipt display/save
- Order Management: search, view line items, status updates, cancellation restricted to ADMIN, archive restricted to ADMIN
- Reports: date filtering, summary totals, export to CSV
- User Accounts (ADMIN): create users, enable/disable, reset passwords
- Settings (ADMIN): system name, currency symbol, receipt footer
- Parameterized OleDb queries and user-facing exception messages

## Requirements

- Windows 10/11
- Visual Studio 2022 with the **.NET desktop development** workload, or the .NET 8 SDK
- Microsoft Access Database Engine (ACE) matching the application's process architecture (prefer x64)
- A blank or newly creatable Access `.accdb` file

> This environment does not contain Windows, Visual Studio/MSBuild, or the ACE provider, so the project has not been compiled or run against Access here. Please build and test on Windows before submitting it for grading.

## Open and run

1. Install Visual Studio 2022 and the .NET desktop development workload.
2. Install the Microsoft Access Database Engine (ACE). Match x64/x86 to your selected build target and existing Office installation.
3. Open `HearthCafe.vbproj` in Visual Studio.
4. Restore NuGet packages and build.
5. Run. The app attempts to create `Data\HearthCafe.accdb` using ADOX if the installed ACE/ADOX components support it. If automatic creation is unavailable, create a blank Access database named `HearthCafe.accdb` inside the generated `Data` folder, then run again. The application creates its tables and indexes on startup.
6. On the first run, select **First-time administrator setup** on the sign-in screen. Create a unique username and a password of at least 8 characters. There is no hardcoded default administrator password.
7. Sign in and create staff accounts from **User Accounts**.

## Database location and backup

The configured path is `Data\HearthCafe.accdb` beside the application. Back up this file regularly. Do not store it in a publicly writable folder. A desktop login does not protect an Access file from a user who can directly open or copy that file; restrict Windows file permissions as well.

## Database schema

The runtime initializes these tables: `Users`, `MenuItems`, `Orders`, `OrderItems`, and `AppSettings`. A reference schema is included in `Database-Schema.sql`. OleDb uses positional `?` parameters, so parameter order in each command is intentional.

## Role permissions

- `ADMIN`: menu maintenance, order management including cancellation/archive, reports, user management, settings.
- `STAFF`: menu viewing, order entry, ordinary order status updates, reports. Staff cannot add/edit/archive menu items, cancel orders, archive orders, manage users, or change settings.

## Rubric checklist (manual verification still required)

- [x] VB.NET Windows Forms project target
- [x] Microsoft Access OleDb data layer and database schema initialization
- [x] CRUD-style menu maintenance (delete is soft archive/restore)
- [x] Login/logout and password hashing
- [x] Role-specific navigation and action restrictions
- [x] Parameterized OleDb commands
- [x] Required-field, numeric, duplicate, discount, and payment validation
- [x] Try/Catch around major user actions and startup
- [x] Database-backed reports with date filters and CSV export
- [ ] Build and run on a Windows machine with ACE installed
- [ ] Live verification of all flows and print/export output on the target machine
- [ ] Confirm the instructor accepts .NET 8 Windows Forms; the rubric specifies Windows Forms but does not specify a framework version

## Important limitations to test before defense

- This is a functional desktop-system baseline, not a substitute for live acceptance testing. Test ACE provider bitness, first-run database creation, all CRUD paths, account disabling/reset, order totals, report date boundaries, and CSV output on the actual lab computer.
- The receipt can be displayed and saved as a text file. Sales reports export to CSV; open the CSV in Excel to print.
- Payment methods are recorded for a café order; this program does not connect to a payment gateway or process card/GCash transactions.
- Access is a file-based database. Apply Windows file permissions and keep backups; application roles alone cannot secure a database file against direct file access.
