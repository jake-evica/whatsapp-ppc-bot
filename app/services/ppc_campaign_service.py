import pandas as pd
from openpyxl import Workbook
from datetime import datetime


class PPCCampaignService:
    def __init__(self, file_path: str, acos_threshold: float = 0.30, match_type: str = "exact", brand_exclusions: list = None):
        self.file_path = file_path
        self.acos_threshold = acos_threshold
        self.match_type = match_type
        self.brand_exclusions = [" " + brand.lower() + " " for brand in (brand_exclusions or [])]
        self.load_data()

    def load_data(self):
        xls = pd.ExcelFile(self.file_path)
        self.df_sponsored = xls.parse('Sponsored Products Campaigns')
        self.df_search_terms = xls.parse('SP Search Term Report')
        self.df_asin_list = xls.parse('ASIN List')

    def filter_keywords(self):
        search_term_col = "Keyword ID"
        orders_col = "Orders"
        acos_col = "ACOS"

        filtered_terms = self.df_search_terms[
            (self.df_search_terms[orders_col] >= 1) &
            (self.df_search_terms[acos_col] < self.acos_threshold)
        ].copy()

        regular_keywords = []
        targeting_keywords = []

        for _, row in filtered_terms.iterrows():
            search_term = str(row[search_term_col]).strip().lower()
            orders = row[orders_col]
            acos = row[acos_col]

            if any(brand in f" {search_term} " for brand in self.brand_exclusions):
                continue

            if search_term.startswith("b0"):  # ASIN targeting
                targeting_keywords.append((search_term, orders, acos))
            else:
                regular_keywords.append((search_term, orders, acos))

        return regular_keywords, targeting_keywords

    def generate_output(self, regular_keywords, targeting_keywords):
        wb = Workbook()
        sheets = {
            "3+ Orders": wb.active,
            "1-2 Orders": wb.create_sheet("1-2 Orders"),
            "Product Targets": wb.create_sheet("Product Targets"),
        }
        sheets["3+ Orders"].title = "3+ Orders"

        headers = ["Keyword", "Orders", "ACOS", "Match Type"]
        for sheet in sheets.values():
            sheet.append(headers)

        for keyword, orders, acos in regular_keywords:
            target_sheet = sheets["3+ Orders"] if orders >= 3 else sheets["1-2 Orders"]
            target_sheet.append([keyword, orders, acos, self.match_type])

        for keyword, orders, acos in targeting_keywords:
            sheets["Product Targets"].append([keyword, orders, acos, "ASIN Targeting"])

        output_file = f"data/processed/amazon_ppc_campaign.xlsx"
        wb.save(output_file)
        return output_file

    def process(self):
        regular_keywords, targeting_keywords = self.filter_keywords()
        output_file = self.generate_output(regular_keywords, targeting_keywords)
        return output_file, regular_keywords