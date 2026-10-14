## 🚀 Auto Loan Securitization & Portfolio Performance Dashboard

✨ **Overview & Executive Summary**

* 📊 The **Auto Loan Securitization and Portfolio Performance Dashboard** is a robust Power BI and DAX-driven analytical solution designed for monitoring, evaluating, and reporting on retail auto loan portfolios, borrower risk profiles, delinquency buckets, and pool performance metrics.
* 📂 The repository encapsulates a complete architecture of data models, dimension tables, automated data quality logs, and advanced DAX measures built for enterprise-grade financial and credit risk dashboards.
* 📈 The analytical framework meticulously tracks critical lending metrics including Exposure at Default, pool factors, delinquency states, non-performing asset classifications, and granular borrower demographics across the securitized pool.

---

🏗️ **Data Model & Dimensional Architecture**

* 🧩 **`DimLoan`**: Captures core loan characteristics including loan identifiers, pool identifiers, servicer names, origination channels, loan purposes, and origination details.
* 👥 **`DimBorrower`**: Provides demographic segmentation encompassing borrower ages, employment types, annual incomes in Indian Rupees, and pre-calculated age/income bands.
* 🗺️ **`DimGeography`**: Establishes geographic hierarchies mapping states to regional portfolios across India.
* 📅 **`DimDate`**: Functions as a full calendar dimension containing years, quarters, month numbers, month names, and fiscal mappings.
* 🏷️ **`DimBucket`**: Defines the delinquency classification framework detailing Days Past Due intervals, 30+ DPD delinquency flags, 90+ NPA flags, and regulatory Reserve Bank of India Special Mention Account classifications.
* ⚙️ **`DimParameters`**: Houses crucial scenario analysis and configuration parameters for stress testing and performance modeling.

---

📊 **Portfolio Composition & Summary Statistics**

* 📋 **Total Sizing**: Comprises **500 individual retail auto loans** representing a total original securitized principal balance of **₹546,073,700.00** (~₹54.61 Crore).
* 🏦 **Servicer Distribution**: Servicing responsibilities are distributed among major institutional partners:
* **Bajaj Finance Auto**: 180 loans (36.0% of the pool)
* **ICICI Auto Loans**: 164 loans (32.8% of the pool)
* **HDFC Bank Auto Finance**: 156 loans (31.2% of the pool)


* 💰 **Borrower Income Analytics**: Borrower annual incomes display a robust distribution:
* **Mean Annual Income**: ₹2,439,622.00
* **Median Annual Income**: ₹2,295,500.00
* **Interquartile Range**: ₹1,289,750.00 (25th percentile) to ₹3,490,000.00 (75th percentile)



---

🧮 **Core DAX Measures & Risk Parameters**

* 📊 **Portfolio Sizing Measures**: Integrates aggregate sums of outstanding and original balances to compute pool factors and total exposure at default.
* 📉 **Volume & Count Measures**: Isolates active performing loans, defaulted loan volumes, and prepayment frequencies.
* ⚠️ **Delinquency Breakdowns**: Segments balances strictly across current statuses, early delinquency buckets, and severe default tiers.
* ⚙️ **Key Securitization Parameters**: Predefined parameters govern performance models:
* **Base Annual CDR**: 3.58%
* **Base Annual CPR**: 9.24%
* **Pool Net LGD (Severity)**: 41.47%
* **Pool Gross LGD**: 65.09%



---

🔍 **Data Governance & Quality Audit**

* 🛡️ **Automated Hygiene Tracking**: Data governance is maintained through automated tracking logs and schema mapping files (`DataQualityLog.csv` and `DataFieldMapping.csv`) that enforce high integrity standards.
* 🛠️ **Anomaly Treatment**: Addresses minor structural anomalies—such as 24 loans carrying zero current balances while retaining historical delinquency statuses—by implementing active loan filters, alongside managing exposure at default variances caused by accrued interest buffers.

---

🚀 **Getting Started & Deployment**

* 📥 **Clone & Access**: Clone the repository files to your local environment and open the interactive Power BI dashboard (`power bi dashboard.pbix`).
* 🔄 **Import DAX Definitions**: Import the centralized DAX measure definitions from `DAXQueryView_AllMeasures.dax` into the model workspace to ensure seamless deployment and reporting accuracy.
