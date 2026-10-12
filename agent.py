import pandas as pd
import tools

try:
    import ollama
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False


class ComplianceDecisionAgent:
    def __init__(self, master_df: pd.DataFrame, use_ollama: bool = False, model_name: str = "llama3.1"):
        self.df = master_df.copy()
        self.use_ollama = use_ollama and OLLAMA_AVAILABLE
        self.model_name = model_name

    def run(self, user_prompt: str) -> str:
        deterministic_answer = self._deterministic_engine(user_prompt)
        if self.use_ollama:
            try:
                system_prompt = (
                    "You are the Nutrition Nest Procurement & Ingredient Compliance AI Agent. "
                    "Always preserve exact numbers, rule IDs, and product names from the Ground Truth data."
                )
                messages = [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Ground Truth Data:\n{deterministic_answer}\n\nQuestion: {user_prompt}"}
                ]
                resp = ollama.chat(model=self.model_name, messages=messages)
                return resp["message"]["content"]
            except Exception as e:
                return f"*(Ollama connection note: `{e}` - showing direct analytical engine output)*\n\n" + deterministic_answer
        return deterministic_answer

    def _deterministic_engine(self, user_prompt: str) -> str:
        prompt_l = user_prompt.lower().strip()

        if any(k in prompt_l for k in ["email", "draft", "outreach", "letter"]):
            matches = tools.tool_lookup_product(self.df, user_prompt)
            if not matches.empty:
                target = matches.iloc[0]
                to_email, email_text = tools.tool_generate_coa_email(target)
                return f"### Generated Vendor Outreach Email\n**Product:** {target['Product name']} ({target['Category']})\n**Recipient:** `{to_email}`\n\n```text\n{email_text}\n```"

        if any(k in prompt_l for k in ["color", "colour", "tier", "l-01", "purple", "orange", "blue", "grey"]):
            if "Priority Color" in self.df.columns:
                c_counts = self.df["Priority Color"].value_counts().to_dict()
                color_desc = {
                    "Red": "L-01 Red (Banned Ingredient / 2+ NOVA-4 Markers)",
                    "Purple": "L-02 Purple (G-03 Undisclosed / Proprietary Formula)",
                    "Orange": "L-03 Orange (G-08 Not Distributed in USA)",
                    "Blue": "L-04 Blue (G-09 Amazon Rating < Threshold)",
                    "Pink": "L-05 Pink (G-07 Unsubstantiated Medical Claim)",
                    "Grey": "L-06 Grey (B-07 Exact Duplicate Row)",
                    "Yellow": "L-07 Yellow (G-01/B-06 Pending Ingredient / Link Data)",
                    "Amber": "L-08 Light Amber (Section 6 Watch-List / COA Hold)",
                    "None": "L-09 Green / No Colour (Passed - Clean)"
                }
                lines = "\n".join([f"- **{color_desc.get(k, k)}**: {v:,} SKUs" for k, v in c_counts.items()])
                return f"### Section 10 Screening Colour Distribution (L-01 to L-09)\n{lines}"

        if any(k in prompt_l for k in ["pass rate", "percentage", "compare", "ratio", "proportion"]):
            total_active = len(self.df)
            passed_total = int((self.df["Screening Status"] == "Passed").sum())
            overall_rate = (passed_total / total_active) * 100 if total_active else 0

            matched_cat = next((c for c in self.df["Category"].unique() if c.lower() in prompt_l), None)
            if matched_cat:
                stats = tools.tool_category_summary(self.df, matched_cat)
                return (
                    f"### Category Compliance Audit: **{stats['Category']}**\n"
                    f"- **Total SKUs Evaluated:** {stats['Total']}\n"
                    f"- **Passed (Clean):** {stats['Passed']} (**{stats['Pass_Rate']:.1f}%**)\n"
                    f"- **Hold (COA / Proof Required):** {stats['Hold']}\n"
                    f"- **Hard Rejected:** {stats['Rejected']}\n"
                    f"- **Pending Data (G-01):** {stats['Pending Data']}"
                )

            rows = []
            for cat in self.df["Category"].unique():
                s = tools.tool_category_summary(self.df, cat)
                if s:
                    rows.append(
                        f"- **{cat}**: {s['Pass_Rate']:.1f}% Passed ({s['Passed']}/{s['Total']}) | "
                        f"Hold: {s['Hold']} | Rejected: {s['Rejected']} | Pending: {s['Pending Data']}"
                    )
            return f"### Multi-Category Compliance Breakdown\n**Overall Store Pass Rate:** **{overall_rate:.1f}%** ({passed_total:,} / {total_active:,} SKUs)\n\n" + "\n".join(rows)

        if any(k in prompt_l for k in ["banned", "toxin", "chemical", "why reject", "top reject", "sweetener", "dye", "nova"]):
            matched_cat = next((c for c in self.df["Category"].unique() if c.lower() in prompt_l), None)
            sub_df = self.df[self.df["Category"] == matched_cat] if matched_cat else self.df
            tox_df = tools.tool_chemical_toxin_summary(sub_df, top_n=8)
            scope = f"in **{matched_cat}**" if matched_cat else "Across Catalog"
            if not tox_df.empty:
                lines = "\n".join([f"- `{r['Screening Rule / Toxin']}` - **{r['Violations Count']} SKUs**" for _, r in tox_df.iterrows()])
                return f"### Top Hard-Gate & Chemical Violations {scope}\n{lines}"

        if any(k in prompt_l for k in ["vendor", "manufacturer", "brand", "supplier"]) and any(
            k in prompt_l for k in ["risk", "worst", "rejection", "fail", "most", "top"]
        ):
            if "Manufacturer name" in self.df.columns:
                rej_df = self.df[self.df["Screening Status"] == "Rejected"]
                bad_mfg = rej_df["Manufacturer name"].dropna().value_counts().head(6).to_dict()
                lines = "\n".join([f"- **{m}**: {cnt} rejected SKUs" for m, cnt in bad_mfg.items()])
                return f"### Vendor Risk Concentration (Highest Rejections)\n{lines}"

        if any(k in prompt_l for k in ["duplicate", "b-07", "variant", "family"]):
            dups = self.df[self.df["Priority Color"] == "Grey"] if "Priority Color" in self.df.columns else pd.DataFrame()
            fam_holds = self.df[self.df["Screening Summary"].astype(str).str.contains("B-07", na=False)]
            return (
                f"### Rule B-07 Audit: Duplicates & Product Variants\n"
                f"- **Exact Duplicate Rows (L-06 Grey - Remove):** {len(dups)} SKUs\n"
                f"- **Total SKUs Flagged for B-07 Variant Grouping / Deduping:** {len(fam_holds)} SKUs"
            )

        if any(k in prompt_l for k in ["coa", "heavy metal", "mycotoxin", "w-06", "glyphosate", "hold", "proof"]):
            hold_df = self.df[self.df["Screening Status"] == "Hold (Proof Required)"]
            counts = hold_df["Category"].value_counts().to_dict()
            lines = "\n".join([f"- **{k}**: {v} SKUs on hold" for k, v in counts.items()])
            return f"### Watch-List & COA Proof Summary ({len(hold_df):,} SKUs on Hold)\n{lines}"

        matches = tools.tool_lookup_product(self.df, user_prompt)
        if not matches.empty:
            out = "### Detailed SKU Compliance Audit\n\n"
            for _, top in matches.head(3).iterrows():
                out += (
                    f"#### **{top['Product name']}** ({top['Category']})\n"
                    f"- **Final Verdict:** `{top['Screening Status']}` | **Section 10 Tier:** `{top.get('Priority Color', 'N/A')}`\n"
                    f"- **Full Screening Issue Code:** `{top['Screening Summary']}`\n"
                    f"- **Manufacturer:** {top.get('Manufacturer name', 'N/A')} | **US Distributed (G-08):** `{top.get('Distributed in USA', 'N/A')}` | **Amazon Rating (G-09):** `{top.get('Product rating (Amazon)', 'N/A')}`\n"
                    f"- **Ingredients on File:** {str(top.get('Ingredients', 'None'))[:280]}\n\n---\n"
                )
            return out

        tot = len(self.df)
        passed = int((self.df["Screening Status"] == "Passed").sum())
        holds = int((self.df["Screening Status"] == "Hold (Proof Required)").sum())
        rej = int((self.df["Screening Status"] == "Rejected").sum())
        pend = int((self.df["Screening Status"] == "Pending Data").sum())

        return (
            f"### Nutrition Nest Compliance Engine Active\n"
            f"Evaluating **{tot:,} SKUs** across **{self.df['Category'].nunique()} active categories**:\n"
            f"- **Passed (No Flags):** {passed:,} ({(passed/tot*100):.1f}%)\n"
            f"- **Hold (Proof / COA Required):** {holds:,} ({(holds/tot*100):.1f}%)\n"
            f"- **Rejected (Hard Gate / Banned):** {rej:,} ({(rej/tot*100):.1f}%)\n"
            f"- **Pending Data (G-01 / B-06):** {pend:,} ({(pend/tot*100):.1f}%)"
        )
