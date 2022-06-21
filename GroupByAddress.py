from difflib import get_close_matches
import pandas as pd

# Read Excel
payMerchantDF = pd.read_excel('./PayMerchantList.xlsx')

# Setup List by Rows
payMerchantAddressList = payMerchantDF["소재지"].values.tolist()

# 상호명: 0, 대표자: 1, 소재지: 2
# Similar strings are defined by a similarity ratio of 0.6 or greater
groups = []
while payMerchantAddressList:
    currAdd = payMerchantAddressList[0]
    similarList = get_close_matches(currAdd, payMerchantAddressList, cutoff=0.8)

    groups.append(similarList)
    
    # Remove addresses in similar list
    for similarAdd in similarList:
        payMerchantAddressList.remove(similarAdd)

# Rebuild Data using Groups
payMerchantMatrix = payMerchantDF.values.tolist()
output = []

for group in groups:
    group.sort()
    for address in group:
        for payMerchant in payMerchantMatrix:
            if (address == payMerchant[2]):
                output.append(payMerchant)
                payMerchantMatrix.remove(payMerchant)
                break

    output.append(["", "", ""])

outputDF = pd.DataFrame(output, columns = ["상호명", "대표자", "소재지"])
outputDF.to_excel("PayMerchantOutput.xlsx", index=False)