#

#  SolarPower太阳能发电 （2）分布式光伏上网电量 其中：①自发自用，余电上网 非自然人  全额上网 其中：自然人 非自然人
# 通用属性
class CommonAttributes:
    # 购电量 powerPurchase 含税 taxIncluded 不含税 excludingTax 机组容量 Unit capacity 供应商名称sap_name 户号 Account
    def __init__(self, name, power_purchase, tax_included, tax_excluding, unit_capacity, sap_name, account):
        self.name = name
        # 购电量 powerPurchase
        self.power_purchase = power_purchase
        # 含税 taxIncluded
        self.tax_included = tax_included
        # 不含税 excludingTax
        self.tax_excluding = tax_excluding
        # 机组容量 Unit capacity
        self.unit_capacity = unit_capacity

        self.sap_name = sap_name
        self.account = account


# 太阳能剩余电量上网-非自然人
class ResidualElectricityNonNaturalPerson(CommonAttributes):
    pass


# 太阳能剩余电量上网-自然人
class ResidualElectricityNaturalPerson(CommonAttributes):
    pass


# 全额上网-自然人
class FullOnlineNaturalPerson(CommonAttributes):
    pass


# 全额上网-非自然人
class FullOnlineNonNaturalPerson(CommonAttributes):
    pass


# 结算信息
class SettlementInformation(CommonAttributes):
    pass


# 水电合计
class HydropowerTotal(CommonAttributes):
    pass


# 农林废弃物
class AgriculturalAndForestryWaste(CommonAttributes):
    pass


# 垃圾焚烧
class WasteIncineration(CommonAttributes):
    pass
