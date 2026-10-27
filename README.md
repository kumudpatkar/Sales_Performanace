# 📊 Sales Performance & Tracking Dashboard

[![Excel](https://shields.io)](https://office.com)
[![VBA](https://shields.io)](https://microsoft.com)

An interactive, corporate-ready **Excel Business Intelligence Dashboard** designed to analyze regional sales performance, monitor individual sales executive quotas, and dynamically toggle between complex analytical layouts using automated VBA backend logic.

---

## 🖼️ Dashboard Preview

![Dashboard Layout](image_1REC3C.png)
> *Tip: Ensure you upload your image file named `image_1REC3C.png` to the root folder of this repository for it to display right here!*

---

## 🚀 Core Features & Architecture

*   **🎛️ Regional Control Matrix:** An integrated master slicer (`Slicer_Region`) spanning across top coordinates to seamlessly filter views for *Central, East, North, South,* and *West* regions.
*   **🔌 Modular Dashboard Toggles:** Custom-built interactive UI check-boxes (`Dashboard 1` to `Dashboard 4`) tied to background boolean flags (`TRUE`/`FALSE`) that actively attach and detach underlying Pivot Caches.
*   **📉 Multidimensional KPI Analysis:**
    *   **Dashboard 1:** Tracks absolute performance volume (`Total Sales` by executive).
    *   **Dashboard 2:** Displays granular breakdown lists for local teams.
    *   **Dashboard 3:** Measures overall project efficiency via `Target Hit %`.
    *   **Dashboard 4:** Highlights execution deficits via structural `Away from Target %` computations.

---

## 💻 Automation & VBA Code Script

The dashboard eliminates system performance lag by utilizing a direct object-oriented approach rather than heavy recorded macros. The macro dynamically monitors true/false conditions across control cells on the sheet layers to map required PivotTable data streams instantly.

```vba
Sub SlicerConnection()

    ' --- PIVOT TABLE 1 (Total Sales View) ---
    If Sheet2.Range("A1").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
            ActiveSheet.PivotTables("PivotTable1"))
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
            ActiveSheet.PivotTables("PivotTable1"))
    End If

    ' --- PIVOT TABLE 2 (Detailed Breakdown) ---
    If Sheet2.Range("A2").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
            ActiveSheet.PivotTables("PivotTable2"))
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
            ActiveSheet.PivotTables("PivotTable2"))
    End If

    ' --- PIVOT TABLE 3 (Target Hit Metrics) ---
    If Sheet2.Range("A3").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
            ActiveSheet.PivotTables("PivotTable3"))
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
            ActiveSheet.PivotTables("PivotTable3"))
    End If

    ' --- PIVOT TABLE 4 (Gap & Variance Analysis) ---
    If Sheet2.Range("A4").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
            ActiveSheet.PivotTables("PivotTable4"))
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
            ActiveSheet.PivotTables("PivotTable4"))
    End If

End Sub
```

---

## 📁 Workbook Layout & Schema

| Sheet Tab Name | Component Category | Functionality |
| :--- | :--- | :--- |
| **`Dashboard`** | User Interface Layer | Contains the primary operational console, slicer elements, toggle controls, and live reporting charts. |
| **`Sheet1`** | Logical Process Layer | Manages background boolean operations, matrix dimensions, and variable registers. |
| **`Raw_Data`** | Database Storage Layer | The normalized transactional data table feeding information directly into the presentation engines. |

---

## 🛠️ Step-by-Step Deployment Guide

1.  **Clone the Repository:** Download the project files onto your desktop local machine.
2.  **Launch Microsoft Excel:** Open the file `sales_performance_dataset.xlsm`.
3.  **Unblock VBA Macros:** Click **"Enable Content" / "Enable Macros"** on the warning security bar to load the connection scripts.
4.  **Explore Data Interactivity:** Use the region filter headers at the top, or check/uncheck dashboards 1-4 to watch the live layout transform instantly.
