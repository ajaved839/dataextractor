from __future__ import annotations

FIELD_COLUMNS = [
    "Source File",
    "MR #",
    "Patient Name",
    "Age",
    "Sex",
    "Contact",
    "Date of Admission",
    "Date of Operation",
    "Date of Discharge",
    "Presenting Complaints",
    "On Lab",
    "Diagnosis",
    "Procedure",
    "Hospital Course",
    "Operative Notes",
    "Follow Up Instructions",
    "Follow Up Time & Date",
    "Discharge Medications",
]

LABEL_ALIASES = {
    "MR #": ["MR #", "MR#", "MR NO", "MR NO.", "MR NUMBER"],
    "Patient Name": ["PATIENT NAME", "NAME"],
    "Age": ["AGE"],
    "Sex": ["SEX", "GENDER"],
    "Contact": ["CONTACT", "CONTACT NO", "CONTACT #", "PHONE", "MOBILE"],
    "Date of Admission": ["DATE OF ADMISSION", "ADMISSION DATE"],
    "Date of Operation": ["DATE OF OPERATION", "OPERATION DATE", "DATE OF SURGERY"],
    "Date of Discharge": ["DATE OF DISCHARGE", "DISCHARGE DATE"],
    "Presenting Complaints": ["PRESENTING COMPLAINTS", "PRESENTING COMPLAINT"],
    "On Lab": ["ON LAB", "LAB", "LABS", "INVESTIGATIONS"],
    "Diagnosis": ["DIAGNOSIS"],
    "Procedure": ["PROCEDURE"],
    "Hospital Course": ["HOSPITAL COURSE"],
    "Operative Notes": ["OPERATIVE NOTES", "OPERATION NOTES", "OPERATIVE NOTE"],
    "Follow Up Instructions": ["FOLLOW UP INSTRUCTIONS", "FOLLOW-UP INSTRUCTIONS"],
    "Follow Up Time & Date": ["FOLLOW UP TIME & DATE", "FOLLOW UP DATE", "FOLLOW-UP DATE"],
    "Discharge Medications": ["DISCHARGE MEDICATIONS", "DISCHARGE MEDICINES", "MEDICATIONS"],
}

SECTION_FIELDS = {
    "Presenting Complaints",
    "On Lab",
    "Diagnosis",
    "Procedure",
    "Hospital Course",
    "Operative Notes",
    "Follow Up Instructions",
    "Follow Up Time & Date",
    "Discharge Medications",
}
