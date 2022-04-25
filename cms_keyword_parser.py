import pandas as pd


class CMSKeywordParser:
    def __init__(self):
        self.file = None
        self.df = None
        self.df_final = None
        self.df_report = None

    def input_file(self, excel_file):
        self.file = excel_file

    def read_file(self):
        self.df = pd.read_excel(
            self.file,
            usecols=["ASIN", "Targeted Search Terms"],
            keep_default_na=False,
            na_values=["_", "NaN"],
        )

    def parse_df(self):
        map = self.df.to_dict("index")
        map[0]["Targeted Search Terms"]
        array = []
        for v in map.values():
            temp_list = []
            if v["Targeted Search Terms"]:
                temp_list.append(v["Targeted Search Terms"].split(","))
            for x in temp_list:
                for y in x:
                    array.append([v["ASIN"], y.strip()])
        for row in array:
            row[1].replace(r"\xa0", " ")
        self.df_final = pd.DataFrame(array, columns=["ASIN", "keyword"])

    def create_report(self):
        df_report = self.df_final.pivot_table(
            index="ASIN", values="keyword", aggfunc="count"
        )
        df_report.reset_index(inplace=True)
        df_report = self.df[["ASIN"]].merge(df_report, how="left")
        df_report["keyword"] = df_report["keyword"].fillna(0)
        self.df_report = df_report

    def export_file(self, outfile_path):
        with pd.ExcelWriter(outfile_path) as writer:
            self.df_final.to_excel(writer, index=False, sheet_name="Parsed Data")
            self.df_report.to_excel(writer, index=False, sheet_name="Report")


def main():
    parser = CMSKeywordParser()
    parser.input_file(
        r"C:\Users\Tin Ha\Documents\CMS Keyword Parser\product-search-results 21st April, 2022 08_59 am.xlsx"
    )
    parser.read_file()
    parser.parse_df()
    parser.create_report()
    parser.export_file("test.xlsx")


if __name__ == "__main__":
    main()
