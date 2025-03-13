import pandas as pd
from typing import Tuple


class PPCBidService:
    @staticmethod
    def _load_data(file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        xls = pd.ExcelFile(file_path)
        ppc_df = pd.read_excel(xls, sheet_name="Sponsored Products")
        asin_df = pd.read_excel(xls, sheet_name="ASIN Data")
        return ppc_df, asin_df

    @staticmethod
    def _prepare_data(ppc_df: pd.DataFrame, asin_df: pd.DataFrame) -> pd.DataFrame:
        ppc_df = ppc_df[[
            "Ad Group ID", "Keyword ID", "Product Targeting ID", "Clicks",
            "Spend", "Sales", "Orders", "ACOS", "CPC", "ROAS"
        ]].copy()

        for col in ["Clicks", "Spend", "Sales", "Orders", "ACOS", "CPC", "ROAS"]:
            ppc_df[col] = pd.to_numeric(ppc_df[col], errors="coerce").fillna(0)

        ppc_df["Product Targeting ID"] = ppc_df["Product Targeting ID"].astype(str)
        asin_df["Product Targeting ID"] = asin_df["ASIN"].astype(str)
        return ppc_df.merge(asin_df, on="Product Targeting ID", how="left")

    @staticmethod
    def _calculate_new_bids(row: pd.Series, target_acos: float) -> float:
        if row["Clicks"] > 0:
            return (row["Sales"] / row["Clicks"]) * target_acos
        return row["CPC"]  # Keep the same bid if there are no clicks

    @staticmethod
    def _adjust_bid(row: pd.Series, target_acos: float) -> float:
        if row["ACOS"] > target_acos:
            return row["New Bid"] * 0.9  # Reduce bid by 10%
        return row["New Bid"] * 1.1  # Increase bid by 10%

    @staticmethod
    def optimize_bids(file_path: str, target_acos: float = 0.3) -> pd.DataFrame:
        ppc_df, asin_df = PPCBidService._load_data(file_path)
        ppc_df = PPCBidService._prepare_data(ppc_df, asin_df)
        ppc_df["New Bid"] = ppc_df.apply(lambda row: PPCBidService._calculate_new_bids(row, target_acos), axis=1)
        ppc_df["Bid Adjustment"] = ppc_df.apply(lambda row: PPCBidService._adjust_bid(row, target_acos), axis=1)
        ppc_df["Final Bid"] = ppc_df["Bid Adjustment"].clip(lower=0.01)
        return ppc_df

    @staticmethod
    def save_results(output_file: str, optimized_df: pd.DataFrame):
        optimized_df.to_excel(output_file, index=False)


if __name__ == "__main__":
    processed_df = PPCBidService.optimize_bids(file_path="data/raw/sample_bid_0001.xlsx")
    PPCBidService.save_results(output_file="data/processed/optimized_bid.xlsx", optimized_df=processed_df)
