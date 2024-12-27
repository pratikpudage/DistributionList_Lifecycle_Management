# Distribution List Lifecycle Management

A project to identify and report distribution list (DL) usage.

## LinkedIn Article

For additional details, read the related article on [LinkedIn](https://www.linkedin.com/article/edit/7278499791079124992/).

---

## How It Works

### 1. Building the Baseline: A Full Inventory of DLs

The first step in managing DL lifecycles is obtaining a comprehensive inventory. This can be challenging, especially in hybrid setups with Exchange Online. To address this, the `Get-DLBaseline.ps1` PowerShell script gathers detailed information for each DL, including:

- **Name of the DL**
- **Primary SMTP Address**
- **Manager's Email Address**
- **Member Count** (to assess size and impact)
- **Status** (active or deleted)

The script outputs the data to an Excel file named `DistributionLists.xlsx`. This baseline provides a snapshot of all DLs, answering critical questions:

- What DLs exist?
- Who manages them?
- How many users depend on them?

Run the `Get-DLBaseline.ps1` script to export this initial baseline.

---

### 2. Tracking Activity: Is the DL Still Being Used?

The next step is determining if a DL is still active. The `Get-DLMessageTrace.ps1` PowerShell script leverages Exchange Online’s message tracing capabilities to track emails sent to each DL over a configurable period (up to 10 days due to Exchange Online limitations).

This script provides insights into active use, as the baseline alone doesn't indicate whether a DL is receiving emails. Message tracking data is stored in a separate file named `MsgTraceDetails.xlsx` for modularity and easier updates.

Run the `Get-DLMessageTrace.ps1` script to export message tracking details.

---

### 3. Merging Insights: A Comprehensive View of DL Usage

The `Process-DLMessageTraceInfo.ps1` script merges the baseline data (`DistributionLists.xlsx`) with message tracking details (`MsgTraceDetails.xlsx`). This process results in a comprehensive report with key insights, including:

- **LastEmailRecdDate**: The most recent email activity for each DL.

The final Excel report provides IT teams with a clear, actionable view of:

- Active DLs
- Dormant DLs
- DLs ready for deletion

This streamlined report supports efficient lifecycle management and helps maintain a clean communication system.

---

## Scripts

- **`Get-DLBaseline.ps1`**: Exports an initial inventory of all DLs.
- **`Get-DLMessageTrace.ps1`**: Tracks email activity for DLs.
- **`Process-DLMessageTraceInfo.ps1`**: Merges baseline and activity data into a final report.

---

## Output Files

1. **`DistributionLists.xlsx`**: Baseline of all DLs.
2. **`MsgTraceDetails.xlsx`**: Email activity data for each DL.
3. **Final Report**: A merged Excel file with a comprehensive overview of DL usage is saved under **`DistributionLists.xlsx`**.
