from typing import Union
from pydantic import BaseModel


class User(BaseModel):
    username: str


class CylinderInfo(BaseModel):
    lubrication: dict
    surface: dict
    deposit: dict
    breakage: dict
    image: str
    remark: str

class PredictionInfo(BaseModel):
    date: str
    total_running_hours: str
    cylinder1: CylinderInfo=None
    cylinder2: CylinderInfo=None
    cylinder3: CylinderInfo=None
    cylinder4: CylinderInfo=None
    cylinder5: CylinderInfo=None
    cylinder6: CylinderInfo=None
    cylinder7: CylinderInfo=None
    cylinder8: CylinderInfo=None
    cylinder9: CylinderInfo=None
    cylinder10: CylinderInfo=None
    cylinder11: CylinderInfo=None
    cylinder12: CylinderInfo=None
    cylinder13: CylinderInfo=None
    cylinder14: CylinderInfo=None
    cylinder15: CylinderInfo=None

class UserInfo(BaseModel):
    _id: str
    email: str
    mobile: str

class Info(BaseModel):
    cylinder_numbers: int
    total_running_hours: str
    running_hrs_since_last: str
    cyl_oil_Type: str
    cyl_oil_consump_Ltr_24hr: str
    normal_service_load_in_percent_MCR: str
    inspection_date: str
    company_name: str
    vessel_name: str
    imo_number: int
    manufacturer: str
    type_of_engine: str
    vessel_type: str

class DocDataModel(BaseModel):
    user: UserInfo
    info: Info
    predictionInfo: PredictionInfo

class UserInput(BaseModel):
    image:str=None
    # cylinder2:str=None
    # cylinder3:str=None
    # cylinder4:str=None
    # cylinder5:str=None
    # cylinder6:str=None
    # cylinder7:str=None
    # cylinder8:str=None
    # cylinder9:str=None
    # cylinder10:str=None
    # cylinder11:str=None
    # cylinder12:str=None
    # cylinder13:str=None
    # cylinder14:str=None
    # cylinder15:str=None

    # VESSEL_OBJECT_ID:int=None
    # EQUIPMENT_CODE:str=None
    # EQUIPMENT_ID:int=None
    # JOB_PLAN_ID:int=None
    # JOB_ID:int=None
    # LOG_ID:int=None

    # #Vessel_Information 
    # Vessel:str=None
    # Hull_No:str=None
    # Vessel_Type:str=None
    # Local_Start_Time:str=None
    # Time_Zone1:str=None
    # Local_End_Time:str=None
    # Time_Zone2:str=None
    # Form_No:str=None
    # IMO_No:str=None

    # #Engine_info
    # Maker:str=None
    # Model:str=None
    # License_Builder:str=None
    # Serial_No:str=None
    # MCR:str=None
    # Speed_at_MCR:str=None
    # Bore:str=None
    # Stroke:str=None

    # #Turbocharger_info
    # Maker_T:str=None
    # Model_T:str=None

    # #General_Data
    # Total_Running_Hour:str=None
    # Cylinder_Oil_type:str=None
    # Normal_service_load_in_percentage_of_MCR:str=None
    # Scrubber:str=None
    # Position: str=None
    # ME_Cyllinder_Oil_Consumption:str=None
    # Me_Running_Hour_Since_Last_Check:str=None
    # Inspected_by_Rank:str=None
    # Fuel_Sulphur_percentage:str=None
