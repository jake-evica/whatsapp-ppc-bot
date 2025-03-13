import uuid
import pandas as pd
import datetime

class AmazonKeywordGenerator:
    @staticmethod
    def load_data(file_path, sheet_name):
        """Load a sheet from an Excel file into a Pandas DataFrame."""
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name).fillna("")
        except ValueError:
            print(f"Error: Sheet '{sheet_name}' not found in the workbook.")
            exit()

    @staticmethod
    def get_sku(df_campaigns, campaign_id):
        """Retrieve the SKU for a given campaign ID."""
        filtered = df_campaigns[(df_campaigns["Campaign ID"] == campaign_id) & (df_campaigns["Entity"] == "Product Ad")]
        skus = filtered["SKU"].unique()

        if len(skus) == 1:
            return skus[0]
        elif len(skus) > 1:
            return "Multi ASIN"
        else:
            return "Not Found"

    @staticmethod
    def safe_float(value):
        """Convert a value to float safely, replacing non-numeric values with 0."""
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    @staticmethod
    def create_campaign(file_path, max_acos = 0.3, excluded_brands = ['Nike', 'Adidas'], match_type = 'Exact'):
        """Generates Amazon keywords based on ACOS threshold and brand exclusions."""
        
        # Load sheets
        df_search = AmazonKeywordGenerator.load_data(file_path, "SP Search Term Report")
        df_campaigns = AmazonKeywordGenerator.load_data(file_path, "Sponsored Products Campaigns")
        df_asin_list = AmazonKeywordGenerator.load_data(file_path, "ASIN List")

        # Convert ASINs to a set for fast lookup
        own_asins = set(df_asin_list["ASIN"].str.upper())

        # Convert excluded brands to lowercase for comparison
        excluded_brands = [brand.lower().strip() for brand in excluded_brands]

        # Get current date
        current_date = datetime.datetime.now().strftime("%Y%m%d")

        # Output lists
        high_orders, high_review = [], []
        low_orders, low_review = [], []
        product_targets, product_targets_review = [], []

        for _, row in df_search.iterrows():
            search_term = row["Customer Search Term"].strip()
            orders = int(row["Orders"])
            acos = AmazonKeywordGenerator.safe_float(row["ACOS"])
            bid = AmazonKeywordGenerator.safe_float(row["Bid"])  # Safe conversion
            campaign_id = row["Campaign ID"]
            ad_group_name = row["Ad Group Name (Informational only)"]

            if orders >= 1 and acos < max_acos:
                exclude = any(brand in search_term.lower() for brand in excluded_brands)

                if not exclude:
                    sku = AmazonKeywordGenerator.get_sku(df_campaigns, campaign_id)
                    keyword_info = [search_term, orders, bid, ad_group_name]

                    if search_term.upper() in own_asins:
                        continue  # Skip our own ASINs

                    if search_term.startswith("b0"):
                        product_targets_review.append(keyword_info)
                    else:
                        if orders >= 3:
                            high_orders.append(keyword_info)
                        else:
                            low_orders.append(keyword_info)

        # Create DataFrame for outputs
        df_high_orders = pd.DataFrame(high_orders, columns=["Keyword", "Orders", "Bid", "Ad Group"])
        df_high_review = pd.DataFrame(high_review, columns=["Keyword", "Orders", "Bid", "Ad Group"])
        df_low_orders = pd.DataFrame(low_orders, columns=["Keyword", "Orders", "Bid", "Ad Group"])
        df_low_review = pd.DataFrame(low_review, columns=["Keyword", "Orders", "Bid", "Ad Group"])
        df_product_targets = pd.DataFrame(product_targets, columns=["Keyword", "Orders", "Bid", "Ad Group"])
        df_product_targets_review = pd.DataFrame(product_targets_review, columns=["Keyword", "Orders", "Bid", "Ad Group"])

        # Save output
        output_file = f"keywords_{str(uuid.uuid4())}.xlsx"
        with pd.ExcelWriter(output_file) as writer:
            df_high_orders.to_excel(writer, sheet_name="3+ Orders", index=False)
            df_high_review.to_excel(writer, sheet_name="3+ Orders - Review", index=False)
            df_low_orders.to_excel(writer, sheet_name="1-2 Orders", index=False)
            df_low_review.to_excel(writer, sheet_name="1-2 Orders - Review", index=False)
            df_product_targets.to_excel(writer, sheet_name="Product Targets", index=False)
            df_product_targets_review.to_excel(writer, sheet_name="Product Targets - Review", index=False)

        return output_file


if __name__ == "__main__":
    file_path = "data/raw/campaign_creation_0001.xlsx"
    output_file = AmazonKeywordGenerator.create_campaign(file_path)
