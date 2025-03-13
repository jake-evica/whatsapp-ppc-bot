import pandas as pd
from typing import List, Tuple


class PPCCampaignService:
    @staticmethod
    def load_data(file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Loads data from an Excel file and returns DataFrames for Sponsored Campaigns, Search Terms, and ASIN List."""
        xls = pd.ExcelFile(file_path)
        df_sponsored = xls.parse('Sponsored Products Campaigns')
        df_search_terms = xls.parse('SP Search Term Report')
        df_asin_list = xls.parse('ASIN List')
        return df_sponsored, df_search_terms, df_asin_list

    @staticmethod
    def filter_keywords(df_search_terms: pd.DataFrame, acos_threshold: float, brand_exclusions: List[str]) -> Tuple[List[Tuple[str, int, float]], List[Tuple[str, int, float]]]:
        """Filters keywords based on ACOS threshold and brand exclusions."""
        search_term_col = "Keyword ID"
        orders_col = "Orders"
        acos_col = "ACOS"

        filtered_terms = df_search_terms[
            (df_search_terms[orders_col] >= 1) &
            (df_search_terms[acos_col] < acos_threshold)
        ].copy()

        regular_keywords = []
        targeting_keywords = []

        brand_exclusions = [" " + brand.lower() + " " for brand in (brand_exclusions or [])]

        for _, row in filtered_terms.iterrows():
            search_term = str(row[search_term_col]).strip().lower()
            orders = row[orders_col]
            acos = row[acos_col]

            if any(brand in f" {search_term} " for brand in brand_exclusions):
                continue

            if search_term.startswith("b0"):  # ASIN targeting
                targeting_keywords.append((search_term, orders, acos))
            else:
                regular_keywords.append((search_term, orders, acos))

        return regular_keywords, targeting_keywords

    @staticmethod
    def generate_dataframe(regular_keywords: List[Tuple[str, int, float]], 
                           targeting_keywords: List[Tuple[str, int, float]], 
                           match_type: str) -> pd.DataFrame:
        """Generates a DataFrame containing campaign data."""
        columns = ["Keyword", "Orders", "ACOS", "Match Type", "Category"]
        data = []

        for keyword, orders, acos in regular_keywords:
            category = "3+ Orders" if orders >= 3 else "1-2 Orders"
            data.append([keyword, orders, acos, match_type, category])

        for keyword, orders, acos in targeting_keywords:
            data.append([keyword, orders, acos, "ASIN Targeting", "Product Targets"])

        df_output = pd.DataFrame(data, columns=columns)
        return df_output

    @staticmethod
    def create_campaign(file_path: str, acos_threshold: float = 0.30, match_type: str = "exact", brand_exclusions: List[str] = None) -> pd.DataFrame:
        """Creates a PPC campaign and returns a DataFrame with processed keywords."""
        df_sponsored, df_search_terms, df_asin_list = PPCCampaignService.load_data(file_path)
        regular_keywords, targeting_keywords = PPCCampaignService.filter_keywords(df_search_terms, acos_threshold, brand_exclusions)
        df_campaign = PPCCampaignService.generate_dataframe(regular_keywords, targeting_keywords, match_type)
        return df_campaign
