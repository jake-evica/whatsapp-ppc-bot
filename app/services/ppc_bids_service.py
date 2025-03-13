import pandas as pd
from typing import Tuple


class PPCBidService:
    def __init__(self, file_path: str, target_acos: float = 0.3):
        self.file_path = file_path
        self.target_acos = target_acos
        self.ppc_df, self.asin_df = self._load_data()

    def _load_data(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        xls = pd.ExcelFile(self.file_path)
        ppc_df = pd.read_excel(xls, sheet_name="Sponsored Products")
        asin_df = pd.read_excel(xls, sheet_name="ASIN Data")
        return ppc_df, asin_df

    def _prepare_data(self) -> pd.DataFrame:
        ppc_df = self.ppc_df[[
            "Ad Group ID", "Keyword ID", "Product Targeting ID", "Clicks", 
            "Spend", "Sales", "Orders", "ACOS", "CPC", "ROAS"
        ]].copy()

        for col in ["Clicks", "Spend", "Sales", "Orders", "ACOS", "CPC", "ROAS"]:
            ppc_df[col] = pd.to_numeric(ppc_df[col], errors="coerce").fillna(0)

        ppc_df["Product Targeting ID"] = ppc_df["Product Targeting ID"].astype(str)
        self.asin_df["Product Targeting ID"] = self.asin_df["ASIN"].astype(str)
        return ppc_df.merge(self.asin_df, on="Product Targeting ID", how="left")

    def _calculate_new_bids(self, row: pd.Series) -> float:
        if row["Clicks"] > 0:
            return (row["Sales"] / row["Clicks"]) * self.target_acos
        return row["CPC"]  # Keep the same bid if there are no clicks

    def _adjust_bid(self, row: pd.Series) -> float:
        if row["ACOS"] > self.target_acos:
            return row["New Bid"] * 0.9  # Reduce bid by 10%
        return row["New Bid"] * 1.1  # Increase bid by 10%

    def optimize_bids(self) -> pd.DataFrame:
        ppc_df = self._prepare_data()
        ppc_df["New Bid"] = ppc_df.apply(self._calculate_new_bids, axis=1)
        ppc_df["Bid Adjustment"] = ppc_df.apply(self._adjust_bid, axis=1)
        ppc_df["Final Bid"] = ppc_df["Bid Adjustment"].clip(lower=0.01)
        return ppc_df

    def save_results(self, output_file: str, optimized_df: pd.DataFrame):
        optimized_df.to_excel(output_file, index=False)