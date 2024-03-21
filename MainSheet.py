from CheckAndPaste import CheckAndPaste
from CopyAndPasteDataMatch import CopyAndPasteDateMatch
from PastHorizontalLine import PasteHorizontalLine
from PastTheData import PastTheData
from PasteFinalResult import PasteFinalResult
from TotalHorizontalLine import TotalHorizontalLine
from VehicleHorizontalLine import VehicleHorizontalLine


def main():
    object = CheckAndPaste("E:\\AMO 20\\sql\\sql27.csv", "E:\\AMO 20\\集計.xlsx")
    object = PastTheData("E:\\AMO 20\\sql\\sql 242.csv", "E:\\AMO 20\\集計.xlsx")
    object = CopyAndPasteDateMatch("E:\\AMO 20\\sql\\sql 360.csv", "E:\\AMO 20\\集計.xlsx")
    object = PasteFinalResult("E:\\AMO 20\\集計.xlsx", "E:\\AMO 20\\LeafTripDistance_20240229.xls")
    object=PasteHorizontalLine("E:\\AMO 20\\LeafTripDistance_20240229.xls","E:\\AMO 20\\LeafTripDistance_Analysis.xlsx")
    object=VehicleHorizontalLine("E:\\AMO 20\\LeafTripDistance_20240229.xls","E:\\AMO 20\\LeafTripDistance_Analysis.xlsx")
    object=TotalHorizontalLine("E:\\AMO 20\\LeafTripDistance_20240229.xls","E:\\AMO 20\\LeafTripDistance_Analysis.xlsx")
main()
