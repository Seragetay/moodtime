import pandas as pd 
import os
import matplotlib as plt
from docx import Document



def read_file():
    # Read the file
    file_data = pd.read_excel(r"C:\Users\E740616\OneDrive - Blue Cross Blue Shield of Michigan\Documents\Projects\CAHPS\Calculator\CHAPS_W_WB.xlsx", sheet_name='Data')
    w_df = file_data[file_data["id_disp"].isin([10, 31])]
    race_filter = ["White", "Black", "Asian", "Hispanic", "North American Native",
                   "Other"]
    disability_filter = ["Unknown", "Y", "N"]
    lis_filter = ["Unknown", "Y", "N"]

    return w_df

############################################################################
wd = read_file()

wd["q45"]
# Assign variables to each category
# Rating_of_Health_Plan question = q38

def single_q(question): # for cat with single question
    valid_col = wd[~wd[question].isin([99, ' '])]
    value = valid_col[question].mean()
    value = round(((value - 0) * 100)/(10-0), 5)

    return value

def care_and_cordnate(question, denm): # for cat with single question
    valid_col = wd[~wd[question].isin([99, ' ', 5, 6, 7])]
    value = valid_col[question].mean()
    value = round(((value - 1) * 100)/(denm-1), 5)

    return value



rating_of_Health_Plan = single_q("q38")
rating_of_Health_Care = single_q("q09")
rating_of_drug_plan = single_q('q45')

# Health Plan Customer Service

getting_information_h = care_and_cordnate("q34", 4)
treated_with_courtesy_respect = care_and_cordnate("q35", 4)
health_plan_forms_easy_to_fill_out = care_and_cordnate("Q36_37", 4)

health_plan_customer_service_lst= [getting_information_h, treated_with_courtesy_respect, health_plan_forms_easy_to_fill_out]
health_plan_customer_service= round(sum(health_plan_customer_service_lst)/len(health_plan_customer_service_lst), 5)

# Getting needed Care
getting_care_tests_or_treatments_necessary = care_and_cordnate("q10", 4)
ease_of_getting_appointment_with_a_specialist = care_and_cordnate("q29", 4)

getting_needed_care_lst = [getting_care_tests_or_treatments_necessary, ease_of_getting_appointment_with_a_specialist]
getting_needed_care = round(sum(getting_needed_care_lst)/len(getting_needed_care_lst), 5)

# Getting Care Quickly
obtaining_needed_care_right_away = care_and_cordnate("q04", 4)
obtaining_care_when_needed = care_and_cordnate("q06", 4)
saw_person_came_to_see_within_15_mins = care_and_cordnate("q08", 4)

getting_care_quickly_lst = [obtaining_needed_care_right_away, obtaining_care_when_needed, saw_person_came_to_see_within_15_mins]
getting_care_quickly = round(sum(getting_care_quickly_lst)/len(getting_care_quickly_lst), 5)

# Getting Needed Prescription Drugs

ease_of_using_drug_plan_to_fill_rx_at_local_pharmacy = care_and_cordnate("q42", 4)
ease_of_using_drug_plan_to_fill_rx_by_mail = care_and_cordnate("q44", 4)
combined_local_pharmacy_and_mail = care_and_cordnate("q42_q44", 4)
ease_of_using_drug_plan_to_get_rx_medicines = care_and_cordnate("q40", 4)

getting_needed_prescription_drugs_lst = [ease_of_using_drug_plan_to_fill_rx_at_local_pharmacy, 
                                        ease_of_using_drug_plan_to_fill_rx_by_mail,
                                        combined_local_pharmacy_and_mail,
                                        ease_of_using_drug_plan_to_get_rx_medicines]

getting_needed_rescription_drugs = round(sum(getting_needed_prescription_drugs_lst)/len(getting_needed_prescription_drugs_lst), 5)

# Care Coordination
combined_item_test_results = care_and_cordnate("q20_q21", 4)
doctor_had_records_info_about_your_care = care_and_cordnate("q18", 4)
doctor_talked_about_prescription_medicines = care_and_cordnate("q23", 4)
got_help_managing_care = 100 - care_and_cordnate("q26", 3)
doctor_informed_and_up_to_date = care_and_cordnate("q32", 4)

care_coordination_lst = [combined_item_test_results,
                        doctor_had_records_info_about_your_care,
                        doctor_talked_about_prescription_medicines,
                        got_help_managing_care,
                        doctor_informed_and_up_to_date]
care_coordination = round(sum(care_coordination_lst)/ len(care_coordination_lst), 5)


Result_dict = {"Rating the health Plan": {"Rating the health Plan": rating_of_Health_Plan},
               "Rating of Healthcare": {"Rating of Health care": rating_of_Health_Care},
               "Rating the drug plan": {"Rating the drug plan": rating_of_drug_plan},
               "Health Plan Customer Service": {"Getting information/help from customer service":getting_information_h,
                                                "Treated with courtesy and respect": treated_with_courtesy_respect,
                                                "Health plan forms easy to fill out": health_plan_forms_easy_to_fill_out,
                                                "Health Plan Customer Service Overall Score": health_plan_customer_service},
                "Getting needed Care": {"Getting care, tests, or treatments necessary": getting_care_tests_or_treatments_necessary,
                                        "Ease of getting appointment with a specialist": ease_of_getting_appointment_with_a_specialist,
                                        "Getting needed Care overall score": getting_needed_care},
               "Getting Care Quickly": {"Obtaining needed care right away": obtaining_needed_care_right_away,
                                        "Obtaining care when needed": obtaining_care_when_needed,
                                        "Saw person came to see within 15 mins": saw_person_came_to_see_within_15_mins,
                                        "Getting Care Quickly overall score": getting_care_quickly},
                "Getting Needed Prescription Drugs": {"Ease of using drug plan to fill rx at local pharmacy": ease_of_using_drug_plan_to_fill_rx_at_local_pharmacy,
                                                      "Ease of using drug plan to fill rx by mail": ease_of_using_drug_plan_to_fill_rx_by_mail,
                                                      "Combined Local Pharmacy and Mail": combined_local_pharmacy_and_mail,
                                                      "Ease of using drug plan to get rx medicines": ease_of_using_drug_plan_to_get_rx_medicines,
                                                      "Getting Needed Prescription Drugs overall score": getting_needed_rescription_drugs},
                "Care Coordination": {"Combined Item - Test Results": combined_item_test_results,
                                      "Doctor had records/info. about your care": doctor_had_records_info_about_your_care,
                                      "Doctor talked about prescription medicines": doctor_talked_about_prescription_medicines,
                                      "Got help managing care": got_help_managing_care,
                                      "Doctor informed and up-to-date": doctor_informed_and_up_to_date,
                                      "Care Coordination": care_coordination}}


indent = "     "
for k, v in Result_dict.items():
    print(k)
    for key, value in v.items():
        print(f"{indent}{key} : {value}")
   # print(f"{k} : {v}")

final_df = pd.DataFrame(Result_dict)



#final_df.to_excel("TESTING.xlsx", index=False)
