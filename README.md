# DistributionList_Lifecycle_Management
Project to identify and report distribution lists usage

# LinkedIn Article
https://www.linkedin.com/article/edit/7278499791079124992/

# How this works

1. Building the Baseline: A Full Inventory of DLs
The first step in managing DL lifecycles is getting a solid understanding of what you’re working with. In many organizations, especially those using hybrid setups with Exchange Online, this is easier said than done. My PowerShell script pulls in a wealth of information for every distribution list:

The name of the DL

The primary SMTP address

The email address of the manager responsible for the DL

The member count (to gauge the size and potential impact of the list)

The status of the DL (whether it’s active or deleted)

This data is captured and stored in a single Excel file called DistributionListsBaseline.xlsx. This baseline serves as the foundation for future comparisons, providing a snapshot of all DLs in the organization. It answers the essential questions: What DLs exist? Who manages them? How many users rely on them?



2. Tracking Activity: Is the DL Still Being Used?
Next comes the critical step of determining whether a DL is still actively used. For this, I created another custom PowerShell script that integrates with Exchange Online’s message tracing capabilities. The script tracks emails sent to each DL over a configurable period (up to the last 10 days due to Exchange Online limitations).

This step is crucial because while the baseline gives you a list of DLs, it doesn’t tell you which ones are still in use. By pulling in message tracking data, we can see which DLs are receiving emails—an indicator of active use. This information is stored in a separate Excel file, MsgTrace.xlsx, to ensure modularity and easier updates in the future.



3. Merging Insights: A Comprehensive View of DL Usage
Finally, the third part of the solution ties everything together. This custom PowerShell script takes the message tracking data from MsgTrace.xlsx and merges it with the baseline from DistributionListsBaseline.xlsx. The result? A complete, up-to-date view of every DL in the system, along with the LastEmailRecdDate field, which shows the most recent email activity for each list.

This gives IT teams a clear, actionable report on which DLs are active, which ones are dormant, and which can potentially be deleted with confidence. The final Excel sheet provides a streamlined way to identify the status of any distribution list, helping you maintain a clean and efficient communication system.
