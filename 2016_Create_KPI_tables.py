__author__ = 'Davisr'

import cx_Oracle
import time
import pyodbc
print "Script Started : "+time.ctime()
import win32com.client
import win32com.client.dynamic
import datetime
import sys
#variables
v_Schema                = 'username'
v_Password              = 'XXXXXXXXX'
v_ODBCConnection        = '@oXXXXX'
v_OracleSignIn          = v_Schema+'/'+v_Password+v_ODBCConnection
v_StartDate             = '1 JAN 2015'
v_EndDate               = '28 DEC 2015'
v_LOS1Year              = '2015'
v_LOS2Year              = '2014'
v_LOS3Year              = '2013'
v_LOS4Year              = '2012'
v_LORestateArray        = []
v_CompTemplateArray     = []
v_FYGDCArray            = []
v_TargetLOS1HiresArray  = []
v_TargetLOS1HiresAndRetention = []
v_AccessDatabaseFile     = 'O:\\MIS\\Score\\Electronic Scoreboard Project\\Copy of WID for new Report Testing.mdb'
v_OutFilePath            = 'O:\\MIS\\Reporting\\Field Bonus Plan Reports\\2016_Comp_Template_Files\\'
v_Connection               = cx_Oracle.connect(v_OracleSignIn)
v_Cursor                   = v_Connection.cursor()
v_ODBC_CONN_STR            = 'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % v_AccessDatabaseFile
v_cnxn                     =  pyodbc.connect(v_ODBC_CONN_STR)
v_MSCursor                 =  v_cnxn.cursor()
v_TargetFilePath           = 'O:\\MIS\\Reporting\\Field Bonus Plan Reports\\2016_Comp_Template_Files\\2016_Comp_Templates_Holding\\'
v_TargetFileName           = '2016SampleTargetsTable.xlsx'
#The target file is provided by Mark Rose.
v_TargetFile               = v_TargetFilePath+v_TargetFileName
#The target file LOS1 Hires for comparison to LOS 1 GT 1000 and for the Composite Retention Rate
# column 57 to 68 (column BE to BP)
v_TargetLOS1Hires        =  57
#The target file location of the employee service number - Column A
v_TargetESN              =  1
# this is the column in the Targets file that shows the amount of Retention Heads the AD or MD LOS 1s needs for that month
v_TargetRetentionHeadCountByMonth = 81
v_Currentyear            = datetime.date.today().strftime("%Y")
v_Currentmonth           = datetime.date.today().strftime("%B")
v_Currentday             = datetime.date.today().strftime("%d")
V_Currenthour            = time.strftime("%X")
v_Currentdate            = v_Currentmonth+"-"+v_Currentday+"-"+v_Currentyear
# v_row                    = 1
# v_QtrColumn              = 2

#Import in the Reporting lo restatement table from MS Acces
v_SQLStatementGetRptingOffice = ("SELECT realignd_lo_cd             ,       "
                                        "TRIM([lo name])            ,       "
                                        "TRIM([reporting lo cd])    ,       "
                                        "TRIM([reporting lo name])  ,       "
                                        "TRIM(region)               ,       "
                                        "TRIM([gm name])                    "
                                        "FROM [table for reporting lo restatment];" )
v_SQLStatementCreate = "CREATE TABLE "+v_Schema+".lo_restate (              " \
                           "realigned_location_code varchar2(14)           , " \
                           "location_name varchar(30)                     , " \
                           "reporting_location_code varchar(14)            , " \
                           "reporting_location_name varchar(30)           , " \
                           "region                  varchar(30)           , " \
                           "managing_director       varchar(30)             " \
                           ")"
v_SQLStatementDrop   = "DROP TABLE "+v_Schema+".lo_restate"
v_SQLStatementInsert = "INSERT INTO "+v_Schema+".lo_restate               ( " \
                           "realigned_location_code                       , " \
                           "location_name                                 , " \
                           "reporting_location_code                       , " \
                           "reporting_location_name                       , " \
                           "region                                        , " \
                           "managing_director                             ) " \
                        "VALUES                                           ( " \
                           ":1, :2, :3, :4, :5, :6                        ) "
#create a table of PAPW for staff codes
v_SQLStatementCreateTbl1_A = ("CREATE TABLE "+v_Schema+".tbl_01_A_TEMP_PAPW_Staff AS "
                                           "SELECT REALIGND_STFF_CD   AS realignd_stff_cd, "
                                           "       COUNT(EMPL_SVC_NUM)AS license_weeks  "
                                           "FROM BD_Schema.OMIMT257_AGT_LINEUP "
                                           "WHERE BSE_ISS_RGST_DT Between '"+v_StartDate+
                                           "' AND '"+v_EndDate+"' "
                                           "AND agt_ctrct_cd IN ('10','14') "
                                           "AND lic_ind ='Y' "
                                           "GROUP BY REALIGND_STFF_CD")
v_SQLStatementDropTbl1_A = ("DROP TABLE  "+v_Schema+".tbl_01_A_TEMP_PAPW_Staff")
v_SQLStatementCreateTbl1_B = ("CREATE TABLE "+v_Schema+".tbl_01_B_TEMP_PAPW_Staff AS "
                                            "SELECT t271.realignd_stff_cd , "
                                                    "SUM(t271.frst_yr_gdc_amt) AS frst_yr_gdc_amt "
                                            "FROM BD_Schema.omimt271_gdc_summary t271 "
                                            "WHERE t271.agt_ctrct_cd IN ('10', '14' ) "
                                            "AND t271.rgst_dt BETWEEN '"+v_StartDate+"' AND '"+v_EndDate+"' "
                                            "GROUP BY t271.realignd_stff_cd ")
v_SQLStatementDropTbl1_B = ("DROP TABLE  "+v_Schema+".tbl_01_B_TEMP_PAPW_Staff")
#following table has to be created after table 3 to get new staff codes
v_SQLStatementCreateTbl1_C = ("CREATE TABLE "+v_Schema+".tbl_01_C_TEMP_PAPW_Staff AS "
                                           "SELECT tbl_1_A.realignd_stff_cd , "
                                                  "tbl_1_A.license_weeks ,  "
                                                  "tbl_1_B.frst_yr_gdc_amt , "
                                                  "tbl_3.latest_stff_cd  "
                                           "FROM tbl_01_A_TEMP_PAPW_Staff tbl_1_A "
                                           "INNER JOIN tbl_01_B_TEMP_PAPW_Staff tbl_1_B ON tbl_1_A.realignd_stff_cd "
                                                      " = tbl_1_B.realignd_stff_cd "
                                           "LEFT JOIN tbl_03_AD_Multiple_Staff_Codes tbl_3 ON tbl_1_A.realignd_stff_cd "
                                                      " = tbl_3.previous_stff_cd " )
v_SQLStatementDropTbl1_C = ("DROP TABLE "+v_Schema+".tbl_01_C_TEMP_PAPW_Staff ")
#Update staff codes that are null so that can combine staffs in a future step.
v_SQLStatementUpdateTbl1_C = ( "UPDATE "+v_Schema+".tbl_01_C_TEMP_PAPW_Staff tbl_1_C "
                                           "SET tbl_1_C.latest_stff_cd = (SELECT tbl_1_C_1.realignd_stff_cd FROM "
                                                                                "tbl_01_C_TEMP_PAPW_Staff tbl_1_C_1 "
                                                                                "WHERE  tbl_1_C_1.realignd_stff_cd "
                                                                                " = tbl_1_C.realignd_stff_cd "
                                                                                " AND tbl_1_C_1.latest_stff_cd IS NULL "
                                                                                ")"
                                           "WHERE tbl_1_C.latest_stff_cd IS NULL ")
v_SQLStatementAlterTbl1_D = ("ALTER TABLE "+v_Schema+".tbl_01_PAPW_by_Staff_Code "
                                           "ADD (ADJ_ID varchar(8), "
                                           "ADJ_COMMENT varchar(30))")
#Insert the records to tbl_01_PAPW_by_Staff_Code from temp table tbl_10_C_TEMP_PAPW_By_Staff_Code
v_SQLStatementInsertTbl1 = ("INSERT INTO "+v_Schema+".tbl_01_PAPW_by_Staff_Code tbl_01 "
                                         "( tbl_01.realignd_stff_cd , "
                                           "tbl_01.license_weeks , "
                                           "frst_yr_gdc_amt , "
                                           "PAPW )"
                                         "SELECT tbl_10_C.reporting_location_code , "
                                                "tbl_10_C.license_weeks , "
                                                "tbl_10_C.frst_yr_gdc_amt , "
                                                "TO_CHAR(tbl_10_C.frst_yr_gdc_amt / tbl_10_C.license_weeks, '$99,999') "
                                         "FROM tbl_10_C_TEMP_PAPW_By_MD tbl_10_C "
                                         "WHERE tbl_10_C.reporting_location_code IS NOT NULL " )
#Create a table with the consolidated staff code and the license weeks and fygdc - will call staff code realignd stff
#code so that it matches other tables
v_SQLStatementCreateTbl1 = ("CREATE TABLE "+v_Schema+".tbl_01_PAPW_by_Staff_Code AS "
                                          "SELECT tbl_1_C.latest_stff_cd AS realignd_stff_cd , "
                                                 "SUM(tbl_1_C.license_weeks) AS license_weeks , "
                                                 "SUM(tbl_1_C.frst_yr_gdc_amt) AS frst_yr_gdc_amt , "
                                                 "TO_CHAR(SUM(tbl_1_C.frst_yr_gdc_amt) / SUM(tbl_1_C.license_weeks) "
                                                 ",  '$99,999' ) AS "
                                                      "PAPW  "
                                          "FROM tbl_01_C_TEMP_PAPW_Staff tbl_1_C "
                                          "GROUP BY tbl_1_C.latest_stff_cd")
v_SQLStatementDropTbl1 = ("DROP TABLE "+v_Schema+".tbl_01_PAPW_by_Staff_Code ")
#Create a table with the FYGDC by staff code
v_SQLStatementCreateTbl2_A = ("CREATE TABLE "+v_Schema+".tbl_02_A_TEMP_FYGDC_by_Staff AS "
                                          "SELECT t271.realignd_stff_cd AS realignd_stff_cd, "
                                          "       SUM(t271.frst_yr_gdc_amt) AS frst_yr_gdc_amt "
                                          "FROM MIM.omimt271_gdc_summary t271 "
                                          "WHERE t271.agt_ctrct_cd IN ('05','07','10','14','15','26','27','51',"
                                          "'60', null) AND t271.rgst_dt BETWEEN '"+v_StartDate+"' AND '"+v_EndDate+"' "
                                          "GROUP BY t271.realignd_stff_cd")
v_SQLStatementDropTbl2_A = ("DROP TABLE  "+v_Schema+".tbl_02_A_TEMP_FYGDC_by_Staff")
v_SQLStatementAlterTbl2_A = ("ALTER TABLE "+v_Schema+".tbl_02_A_TEMP_FYGDC_by_Staff "
                                           "ADD latest_stff_cd varchar(8)")
#update table 2 with the multiple staff latest staff code for the combining of the staffs later in the process
v_SQLStatementUpdateTbl2_A_1 = ("UPDATE "+v_Schema+".tbl_02_A_TEMP_FYGDC_by_Staff tbl_2_A "
                                        "SET tbl_2_A.latest_stff_cd = (SELECT tbl_3.latest_stff_cd "
                                                                    "FROM tbl_03_AD_Multiple_Staff_Codes tbl_3 "
                                                                    "WHERE tbl_2_A.realignd_stff_cd = tbl_3.previous_"
                                                                    "stff_cd) ")
#update table 2 with the those staff codes that only had a single staff code during the year where the latest stff cd
#is null
v_SQLStatementUpdateTbl2_A_2 = ("UPDATE "+v_Schema+".tbl_02_A_TEMP_FYGDC_by_Staff tbl_2_A "
                                        "SET tbl_2_A.latest_stff_cd = (SELECT tbl_2_A_1.realignd_stff_cd "
                                                                      "FROM tbl_02_A_TEMP_FYGDC_by_Staff tbl_2_A_1 "
                                                                      "WHERE tbl_2_A_1.realignd_stff_cd = "
                                                                      "tbl_2_A.realignd_stff_cd) "
                                         "WHERE tbl_2_A.latest_stff_cd IS NULL ")
v_SQLStatementAlterTbl2_B =  ("ALTER TABLE "+v_Schema+".tbl_02_FYGDC_by_Staff_Code "
                                           "ADD (ADJ_ID varchar(8), "
                                           "ADJ_COMMENT varchar(30))")
#Insert MD FYGDC by Reporting Office
v_SQLStatementInsertTbl2 = ("INSERT INTO "+v_Schema+".tbl_02_FYGDC_BY_STAFF_CODE tbl_02 "
                                         "( tbl_02.realignd_stff_cd , "
                                         "  tbl_02.frst_yr_gdc_amt ) "
                                         "SELECT tbl_09_A.reporting_location_code , "
                                                "tbl_09_A.frst_yr_gdc_amt  "
                                         "FROM tbl_09_A_TEMP_FYGDC_BY_MD tbl_09_A " )
#Create table to put the AD FYGDC by Staff Code
v_SQLStatementCreateTbl2 = ("CREATE TABLE "+v_Schema+".tbl_02_FYGDC_by_Staff_Code AS "
                                          "SELECT tbl_2_A.latest_stff_cd AS realignd_stff_cd , "
                                                 "TO_CHAR(SUM(tbl_2_A.frst_yr_gdc_amt), '$99,999,999') "
                                                 "AS frst_yr_gdc_amt "
                                          "FROM tbl_02_A_TEMP_FYGDC_by_Staff tbl_2_A "
                                          "GROUP BY tbl_2_A.latest_stff_cd ")
v_SQLStatementDropTbl2 = ( "DROP TABLE "+v_Schema+".tbl_02_FYGDC_by_Staff_Code ")
#Get a table of staffs and ADs during the time period
v_SQLStatementCreateTbl3_A = ("CREATE TABLE "+v_Schema+".tbl_03_A_TEMP_AD_By_Staff_Code AS "
                                            "SELECT t257.realignd_stff_cd AS realignd_stff_cd ,"
                                                   "t257.empl_svc_num AS empl_svc_num "
                                            "FROM BD_Schema.table_257 t257 "
                                            "WHERE t257.agt_ctrct_cd ='20' "
                                            "AND t257.lic_ind = 'Y' "
                                            "AND t257.bse_iss_rgst_dt BETWEEN '"+v_StartDate+"' AND '"+v_EndDate+"' "
                                            "GROUP BY t257.realignd_stff_cd , "
                                                     "t257.empl_svc_num")
v_SQLStatementDropTbl3_A = ("DROP TABLE  "+v_Schema+".tbl_03_A_TEMP_AD_By_Staff_Code")
#Get a table of those ADs with more than one staff
v_SQLStatementCreateTbl3_B = ("CREATE TABLE "+v_Schema+".tbl_03_B_TEMP_AD_M_Staff_Cd AS "
                                          "SELECT tbl_3.empl_svc_num "
                                          "FROM tbl_03_A_TEMP_AD_By_Staff_Code tbl_3 "
                                          "GROUP BY tbl_3.empl_svc_num "
                                          "HAVING ((COUNT(tbl_3.empl_svc_num)>1))")
v_SQLStatementDropTbl3_B = ("DROP TABLE  "+v_Schema+".tbl_03_B_TEMP_AD_M_Staff_Cd ")
#Get a table of ADs with more than one staff with the different staff codes
v_SQLStatementCreateTbl3_C = ("CREATE TABLE "+v_Schema+".tbl_03_C_TEMP_AD_M_Staff_Cd AS "
                                            "SELECT tbl_3_B.empl_svc_num , "
                                                   "tbl_3_A.realignd_stff_cd "
                                            "FROM tbl_03_A_TEMP_AD_By_Staff_Code tbl_3_A "
                                            "INNER JOIN tbl_03_B_TEMP_AD_M_Staff_Cd tbl_3_B ON "
                                                       "tbl_3_A.empl_svc_num = tbl_3_B.empl_svc_num " )
v_SQLStatementDropTbl3_C = ("DROP TABLE "+v_Schema+".tbl_03_C_TEMP_AD_M_Staff_Cd ")
#Get the ADs with more than one staff the latest staff code
v_SQLStatementCreateTbl3_D = ("CREATE TABLE "+v_Schema+".tbl_03_D_TEMP_AD_M_Staff_Cd AS "
                                          "SELECT tbl_3_B.empl_svc_num , "
                                                 "t257.realignd_stff_cd  "
                                          "FROM BD_Schema.table_257 t257 "
                                          "INNER JOIN tbl_03_B_TEMP_AD_M_Staff_Cd tbl_3_B ON "
                                                     "t257.empl_svc_num = tbl_3_B.empl_svc_num "
                                          "WHERE t257.bse_iss_rgst_dt = '"+v_EndDate+"'")
v_SQLStatementDropTbl3_D = ("DROP TABLE "+v_Schema+".tbl_03_D_TEMP_AD_M_Staff_Cd ")
#Create the Translation or Junction Table with old staff codes and new staff codes to combine metrics
v_SQLStatementCreateTbl3 = ("CREATE TABLE "+v_Schema+".tbl_03_AD_Multiple_Staff_Codes AS "
                                           "SELECT tbl_3_C.realignd_stff_cd AS previous_stff_cd , "
                                                  "tbl_3_C.empl_svc_num , "
                                                  "tbl_3_D.realignd_stff_cd AS latest_stff_cd "
                                           "FROM tbl_03_C_TEMP_AD_M_Staff_Cd tbl_3_C "
                                           "INNER JOIN tbl_03_D_TEMP_AD_M_Staff_Cd tbl_3_D ON "
                                                    "tbl_3_C.empl_svc_num = tbl_3_D.empl_svc_num ")
v_SQLStatementDropTbl3 = ("DROP TABLE "+v_Schema+".tbl_03_AD_Multiple_Staff_Codes ")

#Create a table with the all the staff codes and the individual's first year gdc on that staff
#so if there is an individual on two staffs during the year they will have two records. - this is only LOS 1s
v_SQLStatementCreateTbl4_A = ("CREATE TABLE "+v_Schema+".tbl_04_A_Temp_LOS_1_GT_1000 AS "
                                            "SELECT t271.realignd_stff_cd AS realignd_stff_cd, "
                                                   "t271.empl_svc_num as empl_svc_num        , "
                                                   "SUM(t271.frst_yr_gdc_amt) AS frst_yr_gdc_amt "
                                            "FROM BD_Schema.omimt271_GDC_SUMMARY t271 "
                                            "INNER JOIN BD_Schema.table_041_agent t041 ON t271.empl_svc_num = t041.empl_svc_num "
                                            "WHERE TO_CHAR((t041.emplmt_dt+14), 'YYYY')= '"+v_LOS1Year+"' "
                                            "AND t271.rgst_dt BETWEEN '"+v_StartDate+"' AND '"+v_EndDate+"' "
                                            "GROUP BY t271.realignd_stff_cd,"
                                                     "t271.empl_svc_num" )
v_SQLStatementDropTbl4_A = ("DROP TABLE "+v_Schema+".tbl_04_A_Temp_LOS_1_GT_1000")



#take the data concerning license weeks from the 257 and create a table of LOS1 employee numbers and realigned staff code
#So if the individual is on to staff codes they will have two records. This is only LOS 1 sales agents and staff codes
#during the time period
v_SQLStatementCreateTbl4_B = ("CREATE TABLE "+v_Schema+".tbl_04_B_temp_LOS_1_GT_1000 AS "
                                              "SELECT t257.realignd_stff_cd AS realignd_stff_cd, "
                                                     "t257.empl_svc_num AS empl_svc_num, "
                                                     "COUNT(t257.empl_svc_num) AS license_weeks "
                                              "FROM BD_Schema.table_257 t257 "
                                              "INNER JOIN BD_Schema.table_041_agent t041 ON t257.empl_svc_num = "
                                                         "t041.empl_svc_num "
                                              "WHERE t257.bse_iss_rgst_dt BETWEEN '"+v_StartDate+"' AND '"
                                              +v_EndDate+"' "
                                              "AND TO_CHAR((t041.emplmt_dt+14), 'YYYY') = '"+v_LOS1Year+"' "
                                              "AND t257.agt_ctrct_cd IN ('10','14') "
                                              "AND t257.lic_ind = 'Y' "
                                              "GROUP BY t257.realignd_stff_cd, "
                                                       "t257.empl_svc_num" )
v_SQLStatementDropTbl4_B = ("DROP TABLE "+v_Schema+".tbl_04_B_temp_LOS_1_GT_1000")
#Take the data from the 4C table and update the 4 A table

#Now create a PAPW calculation in the 4A Table by staff and employee

#Creating a temporary table with the realigned staff codes, the number of license weeks, FYGDC and the ESN
#If an individual has two more staff codes during the period they will have more than a single record.
v_SQLStatementCreateTbl4_C = ("CREATE TABLE "+v_Schema+".tbl_04_C_TEMP_LOS_1_GT_1000 AS "
                                          "SELECT tbl_4_A.realignd_stff_cd , "
                                                 "tbl_4_A.empl_svc_num , "
                                                 "tbl_4_B.license_weeks , "
                                                 "tbl_4_A.frst_yr_gdc_amt "
                                          "FROM tbl_04_A_Temp_LOS_1_GT_1000 tbl_4_A "
                                          "LEFT JOIN tbl_04_B_temp_LOS_1_GT_1000 tbl_4_B ON tbl_4_A.empl_svc_num = "
                                                    "tbl_4_B.empl_svc_num "
                                          "AND tbl_4_A.realignd_stff_cd = tbl_4_B.realignd_stff_cd "
                                           )
v_SQLStatementDropTbl4_C = ("DROP TABLE "+v_Schema+".tbl_04_C_TEMP_LOS_1_GT_1000")


v_SQLStatementCreateTbl4_D = ("CREATE TABLE "+v_Schema+".tbl_04_D_TEMP_LOS_1_GT_1000 AS "
                                           "SELECT tbl_3.latest_stff_cd , "
                                                  "tbl_4_C.realignd_stff_cd ,"
                                                  "tbl_4_C.empl_svc_num ,  "
                                                  "tbl_4_C.license_weeks , "
                                                  "tbl_4_C.frst_yr_gdc_amt , "
                                                  "(tbl_4_C.frst_yr_gdc_amt / tbl_4_C.license_weeks) AS PAPW "
                                           "FROM tbl_04_C_TEMP_LOS_1_GT_1000 tbl_4_C "
                                           "LEFT JOIN tbl_03_AD_Multiple_Staff_Codes tbl_3 ON "
                                                       "tbl_4_C.realignd_stff_cd = tbl_3.previous_stff_cd ")


v_SQLStatementDropTbl4_D = ("DROP TABLE "+v_Schema+".tbl_04_D_TEMP_LOS_1_GT_1000 ")
#Update table 4_D and update with latest staff code
v_SQLStatementUpdateTbl4_D_1 = ("UPDATE "+v_Schema+".tbl_04_D_TEMP_LOS_1_GT_1000 table_4 "
                                "SET table_4.latest_stff_cd = (SELECT table_4_d.realignd_stff_cd "
                                                              "FROM tbl_04_D_TEMP_LOS_1_GT_1000 table_4_d "
                                                              "WHERE table_4.realignd_stff_cd = table_4_d.realignd_stff_cd "
                                                              "AND table_4.empl_svc_num = table_4_d.empl_svc_num ) "
                                "WHERE table_4.latest_stff_cd is NULL" )
#Mark wants table tbl_4_E
v_SQLStatementCreateTbl4_E = ("CREATE TABLE "+v_Schema+".tbl_04_E_TEMP_LOS_1_GT_1000 AS "
                                            "SELECT tbl_4_D.latest_stff_cd AS realignd_stff_cd , "
                                                   "tbl_4_D.empl_svc_num , "
                                                   "SUM(tbl_4_D.license_weeks) AS license_weeks , "
                                                   "SUM(tbl_4_D.frst_yr_gdc_amt) AS frst_yr_gdc_amt , "
                                                   "SUM(tbl_4_D.frst_yr_gdc_amt) / SUM(tbl_4_D.license_weeks) AS PAPW "
                                                   "FROM tbl_04_D_TEMP_LOS_1_GT_1000 tbl_4_D "
                                                   "GROUP BY tbl_4_D.empl_svc_num , "
                                                            "tbl_4_D.latest_stff_cd ")

v_SQLStatementDropTbl4_E = ("DROP TABLE "+v_Schema+".tbl_04_E_TEMP_LOS_1_GT_1000")
#Insert the MD data into the table with the AD data
v_SQLStatementInsertTbl4_F_1 = ("INSERT INTO  "+v_Schema+".tbl_04_F_TEMP_LOS_1_GT_1000 tbl_4_F "
                                            "(tbl_4_F.realignd_stff_cd ,  "
                                            "tbl_4_F.LOS1_GT_1000 ) "
                                            "SELECT tbl_11_D.reporting_location_code , "
                                                   "tbl_11_D.LOS1_GT_1000  "
                                            "FROM tbl_11_D_TEMP_LOS_1_GT1000_MD tbl_11_D " )
#Create a temporary table with the Count of LOS 1s greater than 1000 by staff code
v_SQLStatementCreateTbl4_F = ("CREATE TABLE "+v_Schema+".tbl_04_F_TEMP_LOS_1_GT_1000 AS "
                                         "SELECT tbl_4_E.realignd_stff_cd , "
                                                "COUNT(tbl_4_E.empl_svc_num) AS LOS1_GT_1000 "
                                         "FROM tbl_04_E_TEMP_LOS_1_GT_1000 tbl_4_E "
                                         "WHERE tbl_4_E.PAPW >= 1000 "
                                         "GROUP BY tbl_4_E.realignd_stff_cd " )

v_SQLStatementDropTbl4_F = ("DROP TABLE "+v_Schema+".tbl_04_F_TEMP_LOS_1_GT_1000 ")
#Add the AD and MD employee service number and data to the table tbl_04_F_TEMP_LOS_1_GT_1000
v_SQLStatementCreateTbl4_G = ("CREATE TABLE "+v_Schema+".tbl_04_G_TEMP_LOS_1_GT_1000 AS "
                                            "SELECT tbl_6.region , "
                                                   "tbl_6.reporting_location_code , "
                                                   "tbl_6.reporting_location_name , "
                                                   "tbl_6.realignd_stff_cd , "
                                                   "tbl_6.realignd_lo_cd , "
                                                   "tbl_6.agt_frst_name , "
                                                   "tbl_6.agt_lst_name , "
                                                   "tbl_6.empl_svc_num , "
                                                   "tbl_4_F.LOS1_GT_1000 "
                                            "FROM tbl_06_Agency_Directors tbl_6 "
                                            "LEFT JOIN tbl_04_F_TEMP_LOS_1_GT_1000 tbl_4_F ON tbl_6.realignd_stff_cd = "
                                                      "tbl_4_F.realignd_stff_cd " )
v_SQLStatementDropTbl4_G = ("DROP TABLE "+v_Schema+".tbl_04_G_TEMP_LOS_1_GT_1000 ")

v_SQLStatementInsertTbl4_H_1 =("INSERT INTO "+v_Schema+".tbl_04_H_TEMP_LOS_1_GT_1000 ( "
                                            "empl_svc_num , "
                                            "LOS1_Monthly_Hire_Target ) "
                                            "VALUES "
                                              "( :1, :2 )"     )
#Create a temporary table to store the employee service number of the md and ad and the monthly LOS 1 hire target
v_SQLStatementCreateTbl4_H = ("CREATE TABLE "+v_Schema+".tbl_04_H_TEMP_LOS_1_GT_1000 "
                                            "(empl_svc_num varchar2(26) , "
                                            "LOS1_Monthly_Hire_Target varchar2(24) )")
v_SQLStatementDropTbl4_H = ("DROP TABLE "+v_Schema+".tbl_04_H_TEMP_LOS_1_GT_1000")
#Add the LOS 1 Target Hires to the LOS 1 GT 1000 for the ADs and the MDs in a new table
v_SQLStatementCreateTbl4 = ("CREATE TABLE "+v_Schema+".tbl_04_LOS_1_GT_1000 AS "
                                          "SELECT tbl_4_H.empl_svc_num AS TargetFile_ESN , "
                                                 "tbl_4_G.empl_svc_num , "
                                                 "tbl_4_G.region , "
                                                 "tbl_4_G.reporting_location_code , "
                                                 "tbl_4_G.reporting_location_name , "
                                                 "tbl_4_G.realignd_stff_cd , "
                                                 "tbl_4_G.realignd_lo_cd , "
                                                 "tbl_4_G.agt_frst_name , "
                                                 "tbl_4_G.agt_lst_name , "
                                                 "tbl_4_G.los1_gt_1000 , "
                                                 "tbl_4_H.los1_monthly_hire_target , "
                                                 "TO_CHAR(TO_NUMBER(tbl_4_G.los1_gt_1000) / TO_NUMBER(tbl_4_H.los1_monthly_hire_target),'99999.99') AS LOS1GT1000Percentage "
                                          "FROM tbl_04_H_TEMP_LOS_1_GT_1000 tbl_4_H "
                                          "INNER JOIN tbl_04_G_TEMP_LOS_1_GT_1000 tbl_4_G ON tbl_4_H.empl_svc_num = "
                                                     "tbl_4_G.empl_svc_num ")
v_SQLStatementDropTbl4 = ("DROP TABLE "+v_Schema+".tbl_04_LOS_1_GT_1000")
#Insert the Reporting Office Counts into the table tbl_05_Sales_Agents_by_Staff
v_SQLStatementInsertTbl5 = ("INSERT INTO "+v_Schema+".tbl_05_A_TEMP_Sales_Agt tbl_5 "
                                         "(tbl_5.realignd_stff_cd , "
                                         " tbl_5.Num_Sales_Agts )"
                                         "SELECT tbl_12_A.reporting_location_code , "
                                                "tbl_12_A.Num_Sales_Agts "
                                         "FROM tbl_12_A_TEMP_SALES_AGENTS_MD tbl_12_A ")
#Get Sales Agent Count by staff Code
v_SQLStatementAlterTbl5_A = ("ALTER TABLE "+v_Schema+".tbl_05_A_TEMP_Sales_Agt "
                                           "ADD (ADJ_ID varchar(8), "
                                           "ADJ_COMMENT varchar(30))")
v_SQLStatementCreateTbl5_A = ("CREATE TABLE "+v_Schema+".tbl_05_A_TEMP_Sales_Agt AS "
                                           "SELECT t257.realignd_stff_cd AS realignd_stff_cd, "
                                                  "COUNT(t257.empl_svc_num) AS Num_Sales_Agts "
                                           "FROM BD_Schema.table_257 t257 "
                                           "INNER JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num "
                                                                             "= t257.empl_svc_num "
                                           "INNER JOIN BD_Schema.omimt040_agy ON t041.agt_hier_key_num "
                                                                          "= BD_Schema.omimt040_agy.agt_hier_key_num "
                                                                        "AND BD_Schema.omimt040_agy.stff_cd  "
                                                                          "= t257.realignd_stff_cd "
                                           "WHERE t257.agt_ctrct_cd IN ('10','14') "
                                           "AND t257.LIC_IND = 'Y' "
                                           "AND t257.bse_iss_rgst_dt = '"+v_EndDate+"' "
                                           "GROUP BY t257.realignd_stff_cd " )
v_SQLStatementDropTbl5_A = ("DROP TABLE "+v_Schema+".tbl_05_A_TEMP_Sales_Agt ")
#Create a table of the Sales Agent Count with the AD and MD Data
v_SQLStatementCreateTbl5 = ("CREATE TABLE "+v_Schema+".tbl_05_Sales_Agt_Cnt AS "
                                           "SELECT  tbl_6.region , "
                                                   "tbl_6.reporting_location_code , "
                                                   "tbl_6.reporting_location_name , "
                                                   "tbl_6.realignd_stff_cd , "
                                                   "tbl_6.realignd_lo_cd , "
                                                   "tbl_6.agt_frst_name , "
                                                   "tbl_6.agt_lst_name , "
                                                   "tbl_6.empl_svc_num , "
                                                   "tbl_5_A.num_sales_agts , "
                                                   "tbl_5_A.adj_id , "
                                                   "tbl_5_A.adj_comment "
                                           "FROM tbl_06_Agency_Directors tbl_6 "
                                           "LEFT JOIN tbl_05_A_TEMP_SALES_Agt tbl_5_A ON tbl_6.realignd_stff_cd = "
                                                                                        "tbl_5_A.realignd_stff_cd " )

v_SQLStatementDropTbl5 = ("DROP TABLE "+v_Schema+".tbl_05_Sales_Agt_Cnt ")
#Create a list of Managing Directors from temp tables 07
v_SQLStatementInsertTbl6 = ("INSERT INTO "+v_Schema+".tbl_06_Agency_Directors tbl_06 "
                                         "( tbl_06.region ,"
                                          " tbl_06.reporting_location_code , "
                                          " tbl_06.reporting_location_name , "
                                          " tbl_06.realignd_stff_cd , "
                                          " tbl_06.agt_frst_name , "
                                          " tbl_06.agt_mid_name , "
                                          " tbl_06.agt_lst_name , "
                                          " tbl_06.empl_svc_num , "
                                          " tbl_06.emplmt_dt , "
                                          " tbl_06.assm_curr_pstn_dt) "
                                          "SELECT tbl_07_C.region , "
                                                 "tbl_07_C.reporting_location_code , "
                                                 "tbl_07_C.reporting_location_name , "
                                                 "tbl_07_C.reporting_location_code , "
                                                 "tbl_07_C.agt_frst_name , "
                                                 "tbl_07_C.agt_mid_name , "
                                                 "tbl_07_C.agt_lst_name , "
                                                 "CASE WHEN (tbl_07_C.empl_svc_num is NULL) THEN '000000' ELSE "
                                                            "tbl_07_C.empl_svc_num END as empl_svc_num , "
                                                 "tbl_07_C.emplmt_dt , "
                                                 "tbl_07_C.assm_curr_pstn_dt "
                                          "FROM tbl_07_C_TEMP_REPORTING_OFF tbl_07_C "
                               )
#Create a list of Agency Directors and Staff Codes
v_SQLStatementCreateTbl6 = ("CREATE TABLE "+v_Schema+".tbl_06_Agency_Directors AS "
                                          "SELECT lo.region , "
                                                 "lo.reporting_location_code , "
                                                 "lo.reporting_location_name , "
                                                 "t257.realignd_stff_cd , "
                                                 "t257.realignd_lo_cd , "
                                                 "TRIM(t041.agt_frst_name) AS agt_frst_name , "
                                                 "TRIM(t041.agt_mid_name) AS agt_mid_name , "
                                                 "TRIM(t041.agt_lst_name) AS agt_lst_name , "
                                                 "t041.empl_svc_num , "
                                                 "t041.emplmt_dt , "
                                                 "t041.assm_curr_pstn_dt "
                                          "FROM BD_Schema.table_257 t257 "
                                          "INNER JOIN BD_Schema.table_041_agent t041 ON t257.empl_svc_num = t041.empl_svc_num "
                                          "INNER JOIN lo_restate lo ON t257.realignd_lo_cd =lo.realigned_location_code "
                                          "WHERE t041.agt_stat_cd = 'A' "
                                          "AND t041.agt_ctrct_cd = '20' "
                                          "AND lo.reporting_location_code != 'WH01' "
                                          "AND t257.bse_iss_rgst_dt = '"+v_EndDate+"' ")
v_SQLStatementDropTbl6 = ("DROP TABLE "+v_Schema+".tbl_06_Agency_Directors ")

v_SQLStatementCreateTbl7_A = ("CREATE TABLE "+v_Schema+".tbl_07_A_TEMP_REPORTING_OFF AS "
                                            "SELECT restate.region , "
                                                    "restate.reporting_location_code, "
                                                    "restate.reporting_location_name, "
                                                    "restate.realigned_location_code "
                                             "FROM davisr.lo_restate restate "
                                             "WHERE restate.region IN ('SOUTHEASTERN', 'MID AMERICA', 'WESTERN') "
                                             "AND restate.REPORTING_LOCATION_CODE NOT IN ('MA14','WE14','GA87','SE14', "
                                             "'GA75', 'TXL1','TX51') "
                                             "GROUP BY restate.region , "
                                             "restate.reporting_location_code, "
                                             "restate.reporting_location_name, "
                                             "restate.realigned_location_code ")
v_SQLStatementDropTbl7_A = ("DROP TABLE "+v_Schema+".tbl_07_A_TEMP_REPORTING_OFF ")
v_SQLStatementCreateTbl7_B = ("CREATE TABLE "+v_Schema+".tbl_07_B_TEMP_REPORTING_OFF AS "
                                            "SELECT tbl_7_A.region , "
                                                    "tbl_7_A.reporting_location_code, "
                                                    "tbl_7_A.reporting_location_name, "
                                                    "tbl_7_A.realigned_location_code, "
                                                    "TRIM(t041.agt_frst_name) AS agt_frst_name, "
                                                    "TRIM(t041.agt_mid_name) AS agt_mid_name, "
                                                    "TRIM(t041.agt_lst_name) AS agt_lst_name, "
                                                    "t041.empl_svc_num , "
                                                    "t041.emplmt_dt, "
                                                    "t041.assm_curr_pstn_dt "
                                             "FROM "+v_Schema+".tbl_07_A_TEMP_REPORTING_OFF tbl_7_A "
                                             "LEFT JOIN BD_Schema.table_040_staff ON tbl_7_A.realigned_location_code = "
                                                    "t040.lo_cd "
                                             "LEFT JOIN BD_Schema.table_041_agent t041 ON t040.agt_hier_key_num = "
                                                    "t041.agt_hier_key_num "
                                             "WHERE t041.agt_ctrct_cd IN ('30','35','40')"
                                             "AND t041.trmn_dt is null "  )
v_SQLStatementDropTbl7_B = ("DROP TABLE "+v_Schema+".tbl_07_B_TEMP_REPORTING_OFF ")
v_SQLStatementCreateTbl7_C = ("CREATE TABLE "+v_Schema+".tbl_07_C_TEMP_REPORTING_OFF AS "
                                            "SELECT  tbl_7_A.region , "
                                                    "tbl_7_A.reporting_location_code, "
                                                    "tbl_7_A.reporting_location_name, "
                                                    "tbl_7_B.agt_frst_name, "
                                                    "tbl_7_B.agt_mid_name, "
                                                    "tbl_7_B.agt_lst_name, "
                                                    "tbl_7_B.empl_svc_num, "
                                                    "tbl_7_B.emplmt_dt, "
                                                    "tbl_7_B.assm_curr_pstn_dt "
                                            "FROM "+v_Schema+".tbl_07_A_TEMP_REPORTING_OFF tbl_7_A "
                                            "LEFT JOIN "+v_Schema+".tbl_07_B_TEMP_REPORTING_OFF tbl_7_B ON "
                                                       "tbl_7_A.reporting_location_code = "
                                                       "tbl_7_B.reporting_location_code "
                                            "GROUP BY tbl_7_A.region , "
                                                    "tbl_7_A.reporting_location_code, "
                                                    "tbl_7_A.reporting_location_name, "
                                                    "tbl_7_B.agt_frst_name, "
                                                    "tbl_7_B.agt_mid_name, "
                                                    "tbl_7_B.agt_lst_name, "
                                                    "tbl_7_B.empl_svc_num, "
                                                    "tbl_7_B.emplmt_dt, "
                                                    "tbl_7_B.assm_curr_pstn_dt " )
v_SQLStatementDropTbl7_C = ("DROP TABLE "+v_Schema+".tbl_07_C_TEMP_REPORTING_OFF ")
#Calculate the FYGDC for MDs so the data can be placed in the tbl_02_fygdc_by_staff_code
v_SQLStatementCreateTbl9_A = ("CREATE TABLE "+v_Schema+".tbl_09_A_TEMP_FYGDC_By_MD AS "
                                            "SELECT tbl_07_A.reporting_location_code , "
                                                   "SUM(t271.frst_yr_gdc_amt) AS frst_yr_gdc_amt "
                                            "FROM BD_Schema.omimt271_gdc_summary t271 "
                                            "RIGHT JOIN tbl_07_A_TEMP_REPORTING_OFF tbl_07_A ON t271.realignd_lo_cd = "
                                                       "tbl_07_A.realigned_location_code "
                                            "WHERE t271.agt_ctrct_cd IN ('05','07','10','14','15','20','26','27', "
                                                                        "'30', '34', '35', '40', '51', '60' ) "
                                            "AND t271.rgst_dt BETWEEN '"+v_StartDate+"' AND '"+v_EndDate+"' "
                                            "GROUP BY tbl_07_A.reporting_location_code "  )
v_SQLStatementDropTbl9_A = ("DROP TABLE "+v_Schema+".tbl_09_A_TEMP_FYGDC_by_MD ")
#Create a temporary table to start gathering the PAPW by MD
#Start by getting a count of license weeks for the different Reporting Offices for Sales Agents
# Previously when I pulled the data and tested it there were 188 hours associated with a Reporting Office of Null
# these were the following staff codes: XFA10991, SE140991, QM500991
v_SQLStatementCreateTbl10_A = ("CREATE TABLE "+v_Schema+".tbl_10_A_TEMP_PAPW_By_MD AS "
                                             "SELECT tbl_7_A.region , "
                                                    "tbl_7_A.reporting_location_code, "
                                                    "tbl_7_A.reporting_location_name, "
                                                    "COUNT(t257.empl_svc_num) AS license_weeks "
                                             "FROM tbl_07_A_TEMP_REPORTING_OFF tbl_7_A "
                                                   "RIGHT JOIN BD_Schema.table_257 t257 ON t257.realignd_lo_cd  =  "
                                                   "tbl_7_A.realigned_location_code "
                                             "WHERE t257.lic_ind = 'Y' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14') "
                                             "AND BSE_ISS_RGST_DT Between '"+v_StartDate+
                                                                           "' AND '"+v_EndDate+"' "
                                             "GROUP BY tbl_7_A.region , "
                                                    "tbl_7_A.reporting_location_code, "
                                                    "tbl_7_A.reporting_location_name "
                                                     )
v_SQLStatementDropTbl10_A = ("DROP TABLE "+v_Schema+".tbl_10_A_TEMP_PAPW_By_MD ")
#Create a temporary table to start gathering the FYGDC for sales agents for the PAPW by MD
v_SQLStatementCreateTbl10_B = ("CREATE TABLE "+v_Schema+".tbl_10_B_TEMP_PAPW_By_MD AS "
                                             "SELECT tbl_7_A.region , "
                                                    "tbl_7_A.reporting_location_code , "
                                                    "tbl_7_A.reporting_location_name , "
                                                    "SUM(t271.frst_yr_gdc_amt) AS frst_yr_gdc_amt "
                                             "FROM BD_Schema.omimt271_gdc_summary t271 "
                                             "LEFT JOIN tbl_07_A_TEMP_REPORTING_OFF tbl_7_A ON "
                                                       "tbl_7_A.realigned_location_code = t271.realignd_lo_cd "
                                             "WHERE t271.agt_ctrct_cd IN ('10','14') "
                                             "AND t271.rgst_dt BETWEEN '"+v_StartDate+"' AND '"+v_EndDate+"' "
                                             "GROUP BY tbl_7_A.region , "
                                                    "tbl_7_A.reporting_location_code , "
                                                    "tbl_7_A.reporting_location_name  "  )
v_SQLStatementDropTbl10_B = ("DROP TABLE "+v_Schema+".tbl_10_B_TEMP_PAPW_By_MD ")
#Create a temporary table with the License Weeks and the FYGDC by MD together so that it can later be inserted into
#table tbl_01_PAPW_by_Staff_Code
v_SQLStatementCreateTbl10_C =("CREATE TABLE "+v_Schema+".tbl_10_C_TEMP_PAPW_By_MD AS "
                                            "SELECT tbl_10_A.region , "
                                                   "tbl_10_A.reporting_location_code , "
                                                   "tbl_10_A.reporting_location_name , "
                                                   "tbl_10_A.license_weeks , "
                                                   "tbl_10_B.frst_yr_gdc_amt "
                                            "FROM tbl_10_A_TEMP_PAPW_By_MD tbl_10_A "
                                            "LEFT JOIN tbl_10_B_TEMP_PAPW_By_MD tbl_10_B ON tbl_10_A.reporting_location_code = "
                                                       "tbl_10_B.reporting_location_code " )
v_SQLStatementDropTbl10_C = ("DROP TABLE "+v_Schema+".tbl_10_C_TEMP_PAPW_By_MD ")
#Create a temporary table with the LOS 1 s by Reporting Office and First Year GDC
v_SQLStatementCreateTbl11_A = ("CREATE TABLE "+v_Schema+".tbl_11_A_TEMP_LOS_1_GT1000_MD AS "
                                             "SELECT  restate.reporting_location_code, "
                                                   "t271.empl_svc_num as empl_svc_num        , "
                                                   "SUM(t271.frst_yr_gdc_amt) AS frst_yr_gdc_amt "
                                            "FROM BD_Schema.omimt271_GDC_SUMMARY t271 "
                                            "INNER JOIN BD_Schema.table_041_agent t041 ON t271.empl_svc_num = t041.empl_svc_num "
                                            "INNER JOIN BD_Schema.table_040_staff t040 ON t041.agt_hier_key_num = "
                                                                                "t040.agt_hier_key_num "
                                            "INNER JOIN "+v_Schema+".lo_restate restate ON restate.realigned_location_code = "
                                                                                "t040.lo_cd "
                                            "WHERE TO_CHAR((t041.emplmt_dt+14), 'YYYY')= '"+v_LOS1Year+"' "
                                            "AND t271.rgst_dt BETWEEN '"+v_StartDate+"' AND '"+v_EndDate+"' "
                                            "AND t271.agt_ctrct_cd IN ('10','14') "
                                            "GROUP BY restate.reporting_location_code , "
                                                     "t271.empl_svc_num" )
v_SQLStatementDropTbl11_A = ("DROP TABLE "+v_Schema+".tbl_11_A_TEMP_LOS_1_GT1000_MD ")
#Create a temporary table with the LOS 1 s by Reporting Office and License Weeks
v_SQLStatementCreateTbl11_B = ("CREATE TABLE "+v_Schema+".tbl_11_B_TEMP_LOS_1_GT1000_MD AS "
                                              "SELECT restate.reporting_location_code , "
                                                     "t257.empl_svc_num AS empl_svc_num , "
                                                     "COUNT(t257.empl_svc_num) AS license_weeks "
                                              "FROM BD_Schema.table_257 t257 "
                                              "INNER JOIN BD_Schema.table_041_agent t041 ON t257.empl_svc_num = "
                                                         "t041.empl_svc_num "
                                              "INNER JOIN BD_Schema.table_040_staff t040 ON t041.agt_hier_key_num = "
                                                          "t040.agt_hier_key_num "
                                              "INNER JOIN "+v_Schema+".lo_restate restate ON restate.realigned_location_code = "
                                                                                "t040.lo_cd "
                                              "WHERE t257.bse_iss_rgst_dt BETWEEN '"+v_StartDate+"' AND '"
                                              +v_EndDate+"' "
                                              "AND TO_CHAR((t041.emplmt_dt+14), 'YYYY') = '"+v_LOS1Year+"' "
                                              "AND t257.agt_ctrct_cd IN ('10','14') "
                                              "AND t257.lic_ind = 'Y' "
                                              "GROUP BY restate.reporting_location_code , "
                                                       "t257.empl_svc_num" )
v_SQLStatementDropTbl11_B = ("DROP TABLE "+v_Schema+".tbl_11_B_TEMP_LOS_1_GT1000_MD ")
#Creating a Temporary Table with the Counts of LOS 1s by Reporting Offices that shows how many LOS 1s GT 1000
v_SQLStatementCreateTbl11_C = ("CREATE TABLE "+v_Schema+".tbl_11_C_TEMP_LOS_1_GT1000_MD AS "
                                             "SELECT tbl_7_C.region , "
                                                    "tbl_7_C.reporting_location_code , "
                                                    "tbl_7_C.reporting_location_name , "
                                                    "tbl_11_A.empl_svc_num AS esn, "
                                                    "tbl_11_B.empl_svc_num , "
                                                    "tbl_11_A.frst_yr_gdc_amt , "
                                                    "tbl_11_B.license_weeks , "
                                                    "tbl_11_A.frst_yr_gdc_amt / tbl_11_B.license_weeks AS PAPW "
                                             "FROM tbl_07_C_TEMP_REPORTING_OFF tbl_7_C "
                                             "LEFT JOIN tbl_11_A_TEMP_LOS_1_GT1000_MD tbl_11_A ON tbl_7_C.reporting_location_code = "
                                                                               "tbl_11_A.reporting_location_code "
                                             "LEFT JOIN tbl_11_B_TEMP_LOS_1_GT1000_MD tbl_11_B ON tbl_11_B.empl_svc_num = "
                                                                               "tbl_11_A.empl_svc_num "
                               )
v_SQLStatementDropTbl11_C = ("DROP TABLE "+v_Schema+".tbl_11_C_TEMP_LOS_1_GT1000_MD ")
#Summarize table tbl_11_c_TEMP_LOS_1_GT1000_MD by reporting office
v_SQLStatementCreateTbl11_D = ("CREATE TABLE "+v_Schema+".tbl_11_D_TEMP_LOS_1_GT1000_MD AS "
                                             "SELECT tbl_11_C.region , "
                                                    "tbl_11_C.reporting_location_code , "
                                                    "tbl_11_C.reporting_location_name , "
                                                    "COUNT(tbl_11_C.empl_svc_num) AS LOS1_GT_1000 "
                                             "FROM "+v_Schema+".tbl_11_C_TEMP_LOS_1_GT1000_MD tbl_11_C "
                                             "WHERE tbl_11_C.PAPW >= 1000 "
                                             "GROUP BY tbl_11_C.region , "
                                                      "tbl_11_C.reporting_location_code , "
                                                      "tbl_11_C.reporting_location_name ")
v_SQLStatementDropTbl11_D = ("DROP TABLE "+v_Schema+".tbl_11_D_TEMP_LOS_1_GT1000_MD ")
v_SQLStatementCreateTbl12_A = ("CREATE TABLE "+v_Schema+".tbl_12_A_TEMP_SALES_AGENTS_MD AS "
                                                        "SELECT restate.region ,  "
                                                               "restate.reporting_location_code , "
                                                               "restate.reporting_location_name , "
                                                               "COUNT(t257.empl_svc_num) AS Num_Sales_Agts "
                                                        "FROM BD_Schema.table_257 t257 "

                                                        "LEFT JOIN "+v_Schema+".lo_restate restate ON restate.realigned_location_code = "
                                                                                          "t257.realignd_lo_cd "
                                                        "WHERE t257.agt_ctrct_cd IN ('10', '14') "
                                                        "AND t257.LIC_IND = 'Y' "
                                                        "AND t257.bse_iss_rgst_dt = '"+v_EndDate+"' "
                                                        "GROUP BY restate.region ,  "
                                                               "restate.reporting_location_code , "
                                                               "restate.reporting_location_name  "
                                                                )
v_SQLStatementDropTbl12_A = ("DROP TABLE "+v_Schema+".tbl_12_A_TEMP_SALES_AGENTS_MD ")


v_SQLStatementCreateTbl13_A = ("CREATE TABLE "+v_Schema+".tbl_13_A_TEMP_Retention_Rate AS "
                                                        "SELECT t257.realignd_stff_cd  "
                                                        "FROM BD_Schema.table_257 t257 "
                                                        "WHERE t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                                        "AND t257.realignd_stff_cd NOT LIKE 'Q%' "
                                                        "AND t257.realignd_stff_cd NOT LIKE 'WH%' "
                                                        "AND t257.realignd_stff_cd NOT LIKE 'Z%' "
                                                        "AND t257.realignd_stff_cd NOT LIKE 'X%' "
                                                        "GROUP BY t257.realignd_stff_cd ")

v_SQLStatementDropTbl13_A = ("DROP TABLE "+v_Schema+".tbl_13_A_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_B = ("CREATE TABLE "+v_Schema+".tbl_13_B_TEMP_Retention_Rate AS "
                                             "SELECT t257.realignd_stff_cd ,  "
                                                     "COUNT(t257.empl_svc_num) AS LOS_1_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "INNER JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS1Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY t257.realignd_stff_cd ")
v_SQLStatementDropTbl13_B = ("DROP TABLE "+v_Schema+".tbl_13_B_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_C = ("CREATE TABLE "+v_Schema+".tbl_13_C_TEMP_Retention_Rate AS "
                                             "SELECT t257.realignd_stff_cd ,  "
                                                     "COUNT(t257.empl_svc_num) AS LOS_2_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "INNER JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS2Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY t257.realignd_stff_cd ")
v_SQLStatementDropTbl13_C = ("DROP TABLE "+v_Schema+".tbl_13_C_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_D = ("CREATE TABLE "+v_Schema+".tbl_13_D_TEMP_Retention_Rate AS "
                                             "SELECT t257.realignd_stff_cd ,  "
                                                     "COUNT(t257.empl_svc_num) AS LOS_3_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "INNER JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS3Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY t257.realignd_stff_cd ")
v_SQLStatementDropTbl13_D = ("DROP TABLE "+v_Schema+".tbl_13_D_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_E = ("CREATE TABLE "+v_Schema+".tbl_13_E_TEMP_Retention_Rate AS "
                                             "SELECT t257.realignd_stff_cd ,  "
                                                     "COUNT(t257.empl_svc_num) AS LOS_4_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "INNER JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS4Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY t257.realignd_stff_cd ")
v_SQLStatementDropTbl13_E = ("DROP TABLE "+v_Schema+".tbl_13_E_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_F = ("CREATE TABLE "+v_Schema+".tbl_13_F_TEMP_Retention_Rate AS "
                                             "SELECT t257.realignd_stff_cd ,  "
                                                     "COUNT(t257.empl_svc_num) AS LOS_5_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "INNER JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') < '"+v_LOS4Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY t257.realignd_stff_cd ")
v_SQLStatementDropTbl13_F = ("DROP TABLE "+v_Schema+".tbl_13_F_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_G = ("CREATE TABLE "+v_Schema+".tbl_13_G_TEMP_Retention_Rate AS "
                                             "SELECT tbl_3_A.realignd_stff_cd , "
                                                    "tbl_3_B.LOS_1_CNT  , "
                                                    "tbl_3_C.LOS_2_CNT , "
                                                    "tbl_3_D.LOS_3_CNT , "
                                                    "tbl_3_E.LOS_4_CNT , "
                                                    "tbl_3_F.LOS_5_CNT  "
                                             "FROM tbl_13_A_TEMP_Retention_Rate tbl_3_A "
                                             "LEFT JOIN tbl_13_B_TEMP_Retention_Rate tbl_3_B ON tbl_3_A.realignd_stff_cd = "
                                                                                    "tbl_3_B.realignd_stff_cd "
                                             "LEFT JOIN tbl_13_C_TEMP_Retention_Rate tbl_3_C ON tbl_3_A.realignd_stff_cd = "
                                                                                    "tbl_3_C.realignd_stff_cd "
                                             "LEFT JOIN tbl_13_D_TEMP_Retention_Rate tbl_3_D ON tbl_3_A.realignd_stff_cd = "
                                                                                    "tbl_3_D.realignd_stff_cd "
                                             "LEFT JOIN tbl_13_E_TEMP_Retention_Rate tbl_3_E ON tbl_3_A.realignd_stff_cd = "
                                                                                    "tbl_3_E.realignd_stff_cd "
                                             "LEFT JOIN tbl_13_F_TEMP_Retention_Rate tbl_3_F ON tbl_3_A.realignd_stff_cd = "
                                                                                    "tbl_3_F.realignd_stff_cd ")
v_SQLStatementInsertTbl13_G = ("INSERT INTO "+v_Schema+".TBL_13_G_TEMP_RETENTION_RATE tbl_13_G "
                                                       "(tbl_13_G.realignd_stff_cd , "
                                                        "tbl_13_G.LOS_1_CNT , "
                                                        "tbl_13_G.LOS_2_CNT , "
                                                        "tbl_13_G.LOS_3_CNT , "
                                                        "tbl_13_G.LOS_4_CNT , "
                                                        "tbl_13_G.LOS_5_CNT ) "
                                            "SELECT tbl_13_N.reporting_location_code , "
                                                        "tbl_13_N.LOS_1_CNT , "
                                                        "tbl_13_N.LOS_2_CNT , "
                                                        "tbl_13_N.LOS_3_CNT , "
                                                        "tbl_13_N.LOS_4_CNT , "
                                                        "tbl_13_N.LOS_5_CNT "
                                            "FROM TBL_13_N_TEMP_RETENTION_RATE tbl_13_N " )
v_SQLStatementDropTbl13_G = ("DROP TABLE "+v_Schema+".tbl_13_G_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_H = ("CREATE TABLE "+v_Schema+".tbl_13_H_TEMP_Retention_Rate AS "
                                             "SELECT restate.reporting_location_code , "
                                                    "COUNT(t257.empl_svc_num) AS LOS_1_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "RIGHT JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "RIGHT JOIN "+v_Schema+".lo_restate restate ON t257.realignd_lo_cd = "
                                                                                           "restate.realigned_location_code "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS1Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY restate.reporting_location_code  "  )
v_SQLStatementDropTbl13_H = ("DROP TABLE "+v_Schema+".tbl_13_H_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_I = ("CREATE TABLE "+v_Schema+".tbl_13_I_TEMP_Retention_Rate AS "
                                             "SELECT restate.reporting_location_code , "
                                                    "COUNT(t257.empl_svc_num) AS LOS_2_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "RIGHT JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "RIGHT JOIN "+v_Schema+".lo_restate restate ON t257.realignd_lo_cd = "
                                                                                           "restate.realigned_location_code "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS2Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY restate.reporting_location_code  ")
v_SQLStatementDropTbl13_I = ("DROP TABLE "+v_Schema+".tbl_13_I_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_J = ("CREATE TABLE "+v_Schema+".tbl_13_J_TEMP_Retention_Rate AS "
                                             "SELECT restate.reporting_location_code , "
                                                    "COUNT(t257.empl_svc_num) AS LOS_3_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "RIGHT JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "RIGHT JOIN "+v_Schema+".lo_restate restate ON t257.realignd_lo_cd = "
                                                                                           "restate.realigned_location_code "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS3Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY restate.reporting_location_code ")
v_SQLStatementDropTbl13_J = ("DROP TABLE "+v_Schema+".tbl_13_J_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_K = ("CREATE TABLE "+v_Schema+".tbl_13_K_TEMP_Retention_Rate AS "
                                             "SELECT restate.reporting_location_code , "
                                                   "COUNT(t257.empl_svc_num) AS LOS_4_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "RIGHT JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "RIGHT JOIN "+v_Schema+".lo_restate restate ON t257.realignd_lo_cd = "
                                                                                           "restate.realigned_location_code "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') = '"+v_LOS4Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY restate.reporting_location_code " )
v_SQLStatementDropTbl13_K = ("DROP TABLE "+v_Schema+".tbl_13_K_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_L = ("CREATE TABLE "+v_Schema+".tbl_13_L_TEMP_Retention_Rate AS "
                                             "SELECT restate.reporting_location_code , "
                                                     "COUNT(t257.empl_svc_num) AS LOS_5_Cnt  "
                                             "FROM BD_Schema.table_257 t257 "
                                             "RIGHT JOIN BD_Schema.table_041_agent t041 ON t041.empl_svc_num = t257.empl_svc_num "
                                             "RIGHT JOIN "+v_Schema+".lo_restate restate ON t257.realignd_lo_cd = "
                                                                                           "restate.realigned_location_code "
                                             "WHERE TO_CHAR(t041.emplmt_dt+14, 'YYYY') < '"+v_LOS4Year+"' "
                                             "AND t257.BSE_ISS_RGST_DT = '"+v_EndDate+"' "
                                             "AND t257.agt_ctrct_cd IN ('10', '14' )"
                                             "AND t257.lic_ind = 'Y' "
                                             "GROUP BY restate.reporting_location_code ")
v_SQLStatementDropTbl13_L = ("DROP TABLE "+v_Schema+".tbl_13_L_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_M = ("CREATE TABLE "+v_Schema+".tbl_13_M_TEMP_Retention_Rate AS "
                                             "SELECT DISTINCT restate.reporting_location_code  "
                                             "FROM "+v_Schema+".LO_RESTATE  restate " )
v_SQLStatementDropTbl13_M = ("DROP TABLE "+v_Schema+".tbl_13_M_TEMP_Retention_Rate ")
v_SQLStatementCreateTbl13_N = ("CREATE TABLE "+v_Schema+".tbl_13_N_TEMP_Retention_Rate AS "
                                                 "SELECT tbl_13_M.reporting_location_code , "
                                                        " CASE WHEN tbl_13_H.los_1_cnt IS NULL THEN"
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_H.los_1_cnt "
                                                        " END AS los_1_cnt ,  "
                                                        " CASE WHEN tbl_13_I.los_2_cnt IS NULL THEN"
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_I.los_2_cnt "
                                                        " END AS los_2_cnt ,  "
                                                        " CASE WHEN tbl_13_J.los_3_cnt IS NULL THEN"
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_J.los_3_cnt "
                                                        " END AS los_3_cnt ,  "
                                                        " CASE WHEN tbl_13_K.los_4_cnt IS NULL THEN"
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_K.los_4_cnt "
                                                        " END AS los_4_cnt ,  "
                                                        " CASE WHEN tbl_13_L.los_5_cnt IS NULL THEN"
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_L.los_5_cnt "
                                                        " END AS los_5_cnt   "

                                                 "FROM tbl_13_M_TEMP_Retention_Rate tbl_13_M "
                                                 "LEFT JOIN tbl_13_H_TEMP_Retention_Rate tbl_13_H ON tbl_13_M.reporting_location_code = "
                                                                                                  " tbl_13_H.reporting_location_code "
                                                 "LEFT JOIN tbl_13_I_TEMP_Retention_Rate tbl_13_I ON tbl_13_M.reporting_location_code = "
                                                                                                  " tbl_13_I.reporting_location_code "
                                                 "LEFT JOIN tbl_13_J_TEMP_Retention_Rate tbl_13_J ON tbl_13_M.reporting_location_code = "
                                                                                                  " tbl_13_J.reporting_location_code "
                                                 "LEFT JOIN tbl_13_K_TEMP_Retention_Rate tbl_13_K ON tbl_13_M.reporting_location_code = "
                                                                                                  " tbl_13_K.reporting_location_code "
                                                 "LEFT JOIN tbl_13_L_TEMP_Retention_Rate tbl_13_L ON tbl_13_M.reporting_location_code = "
                                                                                                  " tbl_13_L.reporting_location_code " )
v_SQLStatementDropTbl13_N = ("DROP TABLE "+v_Schema+".tbl_13_N_TEMP_Retention_Rate ")
#Create a table 13_P_TEMP_Retention_Rate with the target hires for the month and the target LOS 1 for the calculation of Retention Rate
v_SQLStatementCreateTbl13_P = ("CREATE TABLE "+v_Schema+".tbl_13_P_TEMP_Retention_Rate  "
                                             "( empl_svc_num varchar2(26) , "
                                               "LOS1_Monthly_Hire_Target varchar2(24) , "
                                               "LOS_Monthly_Retention_Target varchar2(24))" )

v_SQLStatementDropTbl13_P = ("DROP TABLE "+v_Schema+".tbl_13_P_TEMP_Retention_Rate ")
#Insert the data into table 13_P_TEMP_Retention_Rate
v_SQLStatementInsertTbl13_P_1 = ("INSERT INTO "+v_Schema+".tbl_13_P_TEMP_Retention_Rate ( "
                                              "empl_svc_num , "
                                              "LOS1_Monthly_Hire_Target , "
                                              "LOS_Monthly_Retention_Target ) "
                                              "VALUES "
                                              "(:1, :2, :3 )"
                                  )
#Create Table with all the detail
v_SQLStatementCreateTbl13_Q = ( "CREATE TABLE "+v_Schema+".tbl_13_Q_TEMP_Retention_Rate AS "
                                        "SELECT tbl_6.region , "
                                                        " tbl_6.reporting_location_code , "
                                                        " tbl_6.realignd_stff_cd , "
                                                        " tbl_6.realignd_lo_cd , "
                                                        " tbl_6.agt_frst_name , "
                                                        " tbl_6.agt_mid_name , "
                                                        " tbl_6.agt_lst_name , "
                                                        " tbl_6.empl_svc_num , "
                                                        " CASE WHEN tbl_13_G.los_1_cnt IS NULL THEN "
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_G.los_1_cnt "
                                                        " END AS los_1_cnt  , "
                                                        " CASE WHEN tbl_13_G.los_2_cnt IS NULL THEN "
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_G.los_2_cnt "
                                                        " END AS los_2_cnt  , "
                                                        " CASE WHEN tbl_13_G.los_3_cnt IS NULL THEN "
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_G.los_3_cnt "
                                                        " END AS los_3_cnt  , "
                                                        " CASE WHEN tbl_13_G.los_4_cnt IS NULL THEN "
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_G.los_4_cnt "
                                                        " END AS los_4_cnt  , "
                                                        " CASE WHEN tbl_13_G.los_5_cnt IS NULL THEN "
                                                        "           0 "
                                                        " ELSE "
                                                        "         tbl_13_G.los_5_cnt "
                                                        " END AS los_5_cnt  , "
                                                        " tbl_13_P.los1_monthly_hire_target , "
                                                        " tbl_13_P.los_monthly_retention_target  "
                                         "FROM tbl_13_G_TEMP_Retention_Rate tbl_13_G "
                                         "INNER JOIN tbl_06_Agency_Directors tbl_6 ON tbl_13_G.realignd_stff_cd = tbl_6.realignd_stff_cd "
                                         "INNER JOIN tbl_13_P_TEMP_Retention_Rate tbl_13_P ON tbl_13_P.empl_svc_num = tbl_6.empl_svc_num ")
v_SQLStatementDropTbl13_Q = ("DROP TABLE "+v_Schema+".tbl_13_Q_Retention_Rate ")
#Get the Database Table for Reporting Office
v_MSCursor.execute(v_SQLStatementGetRptingOffice)

#read values from MS Access into an array
for row in v_MSCursor.fetchall():
        v_LORestateArray.append([row[0], row[1], row [2], row [3], row[4], row[5]])

#Create Table if already exists drop it and then create it
try:
    v_Cursor.execute(v_SQLStatementCreate)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDrop)
    v_Cursor.execute(v_SQLStatementCreate)
    v_Connection.commit()
print "1. Creating the LO_Restate table with Reporting Offices "
#Now insert the Reporting Data in the newly created table
v_Cursor.executemany(v_SQLStatementInsert, v_LORestateArray)
v_Connection.commit()

print "2. Creating temporary table with Sales Agent License Weeks by Staff Code"

# #create table tbl_01_PAPW_by_StaffCode - rest of construction of table 01 is near line 445 because it uses parts
#of table 03
try:
    v_Cursor.execute(v_SQLStatementCreateTbl1_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl1_A)
    v_Cursor.execute(v_SQLStatementCreateTbl1_A)
    v_Connection.commit()
print "3. Creating temporary table with Sales Agents FYGDC Data by Staff Code"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl1_B)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl1_B)
    v_Cursor.execute(v_SQLStatementCreateTbl1_B)
    v_Connection.commit()
print "4. Create a Temporary Table with the Staff Codes on the End Date \n" \
      "      tbl_13_A_TEMP_Retention_Rate \n" \
      "      from line 1292 moved for performance issue"
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_A)
    v_Connection.commit()
except:
    time.sleep(30)
    v_Cursor.execute(v_SQLStatementDropTbl13_A)
    v_Cursor.execute(v_SQLStatementCreateTbl13_A)
    v_Connection.commit()
print "5. Create a temporary Table with the Agency Directors and their Staff Codes between the start and end dates"
#create table tbl_02_FYGDC_by_StaffCode - creation of the combined staffs is down near line 452 because it uses parts
#of table 03
try:
    v_Cursor.execute(v_SQLStatementCreateTbl2_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl2_A)
    v_Cursor.execute(v_SQLStatementCreateTbl2_A)
    v_Connection.commit()
print "6. Alter the temporary table just built to add the latest staff code field"
v_Cursor.execute(v_SQLStatementAlterTbl2_A)
v_Connection.commit()
#start the process to create tbl_03_AD_Multiple_Staffs so that we can combine those with multiple staff codes
print "7. Create a Temporary Table with the LOS 1 Sales Agents by Staff Code on the End Date \n" \
      "      tbl_13_B_TEMP_Retention_Rate \n"  \
      "      Moved from line 1308 for performance reasons. "
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_B)
    v_Connection.commit()
except:
    time.sleep(30)
    v_Cursor.execute(v_SQLStatementDropTbl13_B)
    v_Cursor.execute(v_SQLStatementCreateTbl13_B)
    v_Connection.commit()
time.sleep(60)

print "8. Create a Temporary Table with the LOS 1 Sales Agents by Reporting Office on the End Date \n" \
      "      tbl_13_H_TEMP_Retention_Rate \n" \
      "      Moved from Line 1372 due to performance issue."
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_H)
    v_Connection.commit()

except:

    v_Cursor.execute(v_SQLStatementDropTbl13_H)
    v_Cursor.execute(v_SQLStatementCreateTbl13_H)
    v_Connection.commit()
time.sleep(60)


print "9. Creating temporary table with the Agency Directors between dates that is grouped by staff code and \n " \
      "      the employee service number between a start and end date \n" \
      "      tbl_03_A_TEMP_AD_M_Staff_Cd"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl3_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl3_A)
    v_Cursor.execute(v_SQLStatementCreateTbl3_A)
    v_Connection.commit()
print "10. Create a temporary table from the above table that has just those Agency Directors that have more than \n " \
      "       a single staff code between the start and stop dates \n "\
      "       tbl_03_B_TEMP_AD_M_Staff_Cd"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl3_B)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl3_B)
    v_Cursor.execute(v_SQLStatementCreateTbl3_B)
    v_Connection.commit()
print "11. Creating temporary table with the Agency Directors with Multiple Staffs and with the multiple staff codes \n"\
      "       tbl_03_C_TEMP_AD_M_Staff_Cd"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl3_C)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl3_C)
    v_Cursor.execute(v_SQLStatementCreateTbl3_C)
    v_Connection.commit()

print "12. Create a Temporary Table with the LOS 2 Sales Agents by Staff Code on the End Date \n" \
      "       tbl_13_C_TEMP_Retention_Rate \n"\
      "       Moved from line 1321 for performance issues. "
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_C)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_C)
    v_Cursor.execute(v_SQLStatementCreateTbl13_C)
    v_Connection.commit()
time.sleep(60)
print "13. Creating temporary table to get a list of the Agency Directors with Multiple Staffs' latest staff code \n " \
      "       on the end date of the time period"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl3_D)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl3_D)
    v_Cursor.execute(v_SQLStatementCreateTbl3_D)
    v_Connection.commit()
print "14. Creating Permanent Table called tbl_03_AD_Multiple_Staff_Codes that has the Agency Directors' staff codes \n" \
      "       then the employee service number and then the latest staff codes. "
try:
    v_Cursor.execute(v_SQLStatementCreateTbl3)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl3)
    v_Cursor.execute(v_SQLStatementCreateTbl3)
    v_Connection.commit()
#create table tbl_04_A_Temp_LOS_1_GT_1000
print "15. Creating Temporary Table tbl_04_A_LOS_1_GT_1000 with the first year gdc and the staff code and \n" \
      "       employee service number for the LOS 1 sales agents between a start and end date if they are \n" \
      "       active in the 041 Table."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4_A)
    v_Cursor.execute(v_SQLStatementCreateTbl4_A)
    v_Connection.commit()
print "16. Create a Temporary Table with the LOS 3 Sales Agents by Staff Code on the End Date \n" \
      "       tbl_13_D_TEMP_Retention_Rate \n" \
      "       Moved from line 1351 due to performance issues. "
try:
    time.sleep(60)
    v_Cursor.execute(v_SQLStatementCreateTbl13_D)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_D)
    v_Cursor.execute(v_SQLStatementCreateTbl13_D)
    v_Connection.commit()
time.sleep(60)
print "17. Creating Temporary Table tbl_04_B_LOS_1_GT_1000 with license weeks from the t257 table between \n" \
      "       a start and end date with for LOS 1 sales agents and that are licensed for that week."
#Create a table with gdc by employee and staff code
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4_B)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4_B)
    v_Cursor.execute(v_SQLStatementCreateTbl4_B)
    v_Connection.commit()
print "18. Creating Temporary Table tbl_04_C_LOS_1_GT_1000 with the realigned staff codes, employee \n" \
      "       service numbers, license weeks and FYGDC. If an agent has been on more than one staff code \n" \
      "       they will appear multiple times in this table."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4_C)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4_C)
    v_Cursor.execute(v_SQLStatementCreateTbl4_C)
    v_Connection.commit()
print "19. Creating Temporary Table tbl_04_D_LOS_1_GT_1000 with the latest staff code for a sales agent, \n" \
      "       their realigned staff code, employee service number, license weeks, first year gdc amount and the \n" \
      "       PAPW. If an individual has been on multiple staffs they will appear multiple times in this table."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4_D)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4_D)
    v_Cursor.execute(v_SQLStatementCreateTbl4_D)
    v_Connection.commit()
print "20. Create a Temporary Table with the LOS 5 Sales Agents by Staff Code on the End Date \n"\
      "       tbl_13_F_TEMP_Retention_Rate\n"\
      "       Moved from line 1358 due to performance issues. "
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_F)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_F)
    v_Cursor.execute(v_SQLStatementCreateTbl13_F)
    v_Connection.commit()
time.sleep(60)
print "21. Update Temporary Table tbl_04_D_LOS_1_GT_1000 field realigned staff code (latest staff code) with the \n" \
      "       realigned staff code / one staff code they have been on if they only have one staff code so far."
v_Cursor.execute(v_SQLStatementUpdateTbl4_D_1)
v_Connection.commit()
print "22. Creating Temporary Table tbl_04_E_LOS_1_GT_1000 of the latest staff code as realignd staff code, employee \n" \
      "       service number, license weeks and frst_yr_gdc_amt and PAPW."
try:
   v_Cursor.execute(v_SQLStatementCreateTbl4_E)
   v_Connection.commit()
except:
   v_Cursor.execute(v_SQLStatementDropTbl4_E)
   v_Cursor.execute(v_SQLStatementCreateTbl4_E)
   v_Connection.commit()
print "23. Creating Permanent Table tbl_04_LOS_1_GT_1000 with the realigned staff code / latest staff code \n" \
      "       and the number of LOS 1 sales agents above 1000 FYGDC."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4_F)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4_F)
    v_Cursor.execute(v_SQLStatementCreateTbl4_F)
    v_Connection.commit()
print "24. Create a Temporary Table with the LOS 4 Sales Agents by Staff Code on the End Date \n" \
      "       tbl_13_E_TEMP_Retention_Rate \n" \
      "       Moved from line 1346 for performace issues."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_E)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_E)
    v_Cursor.execute(v_SQLStatementCreateTbl13_E)
    v_Connection.commit()
time.sleep(60)
#create table 5 with the current staff code and the count of sales agents
print "25. Creating Permant Table tbl_05_Sales_Agents_by_Staff with the number of sales agents on the end date. \n" \
      "       tbl_05_A_TEMP_Sales_A"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl5_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl5_A)
    v_Cursor.execute(v_SQLStatementCreateTbl5_A)
    v_Connection.commit()
#create table 6 with a list of Agency Directors and administrative type data
print "26. Creating Permanent Table tbl_06_Agency_Directors on the end date."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl6)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl6)
    v_Cursor.execute(v_SQLStatementCreateTbl6)
    v_Connection.commit()
print "27. Create a Temporary Table with the LOS 2 Sales Agents by Reporting Office on the End Date \n" \
      "       tbl_13_I_TEMP_Retention_Rate\n" \
      "       Moved from line 1439 due to performance issues."
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_I)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_I)
    v_Cursor.execute(v_SQLStatementCreateTbl13_I)
    v_Connection.commit()
#Combine multiple staffs for PAPW for table 1 uses table 3
try:
    v_Cursor.execute(v_SQLStatementCreateTbl1_C)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl1_C)
    v_Cursor.execute(v_SQLStatementCreateTbl1_C)
    v_Connection.commit()
#update the new staff code for consolidate staff codes or multiple staff codes in tbl_1_C
v_Cursor.execute(v_SQLStatementUpdateTbl1_C)
v_Connection.commit()
# create a final table of PAPW with the new staff code
try:
    v_Cursor.execute(v_SQLStatementCreateTbl1)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl1)
    v_Cursor.execute(v_SQLStatementCreateTbl1)
    v_Connection.commit()
#update table 2 with the multiple staff latest staff code for the combining of the staffs later in the process
v_Cursor.execute(v_SQLStatementUpdateTbl2_A_1)
#update table 2 with the those staff codes that only had a single staff code during the year where the latest stff cd
#is null
v_Cursor.execute(v_SQLStatementUpdateTbl2_A_2)
v_Connection.commit()
#Create table 2 with the new consolidated staff code and the fygdc consolidate
try:
    v_Cursor.execute(v_SQLStatementCreateTbl2)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl2)
    v_Cursor.execute(v_SQLStatementCreateTbl2)
    v_Connection.commit()
print 'Finished Creating All the Tables'
#clean up temporary tables
print "28. Deleting Temporary Tables"
v_Cursor.execute(v_SQLStatementDropTbl1_A)
v_Cursor.execute(v_SQLStatementDropTbl1_B)
v_Cursor.execute(v_SQLStatementDropTbl1_C)
v_Cursor.execute(v_SQLStatementDropTbl2_A)
v_Cursor.execute(v_SQLStatementDropTbl3_A)
v_Cursor.execute(v_SQLStatementDropTbl3_B)
v_Cursor.execute(v_SQLStatementDropTbl3_C)
v_Cursor.execute(v_SQLStatementDropTbl3_D)
v_Cursor.execute(v_SQLStatementDropTbl4_A)
v_Cursor.execute(v_SQLStatementDropTbl4_B)
v_Cursor.execute(v_SQLStatementDropTbl4_C)
#Add Adjustment of number and comment
print "29. Adding the Adjustment number and Adjustment comment to the following tables: \n" \
      "      A) tbl_01_PAPW_By_Staff_Code  \n " \
      "     B) tbl_02_FYGDC_By_Staff_Code \n " \
      "     C) tbl_05_Sales_Agents_by_Staff "
v_Cursor.execute(v_SQLStatementAlterTbl1_D)
v_Cursor.execute(v_SQLStatementAlterTbl2_B)
v_Cursor.execute(v_SQLStatementAlterTbl5_A)

#Start putting the MD Data in the same tables as the ADs

print "30. Creating Temporary Table tbl_07_A_TEMP_Reporting_Off with the region, reporting office code,\n"\
      "       reporting office name, realigned code."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl7_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl7_A)
    v_Cursor.execute(v_SQLStatementCreateTbl7_A)
    v_Connection.commit()
print "31. Creating Temporary Table tbl_07_B_TEMP_Reporting_Off with the Managing Director Information. "
try:
    v_Cursor.execute(v_SQLStatementCreateTbl7_B)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl7_B)
    v_Cursor.execute(v_SQLStatementCreateTbl7_B)
    v_Connection.commit()
print "32. Creating Temporary Table tbl_07_C_TEMP_Reporting_Off with the Managing Director Information with \n" \
      "       the Reporting Offices without a Managers."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl7_C)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl7_C)
    v_Cursor.execute(v_SQLStatementCreateTbl7_C)
    v_Connection.commit()
print "33. Inserting the Managing Directors into the table of Agency Directors - tbl_06_Agency_Directors."
v_Cursor.execute(v_SQLStatementInsertTbl6)
v_Connection.commit()
#Create a temporary table with FYGDC by Reporting Office
print "34. Creating a Temporary Table with the first year gdc by reporting office, tbl_09_A_TEMP_FYGDC_By_MD."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl9_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl9_A)
    v_Cursor.execute(v_SQLStatementCreateTbl9_A)
    v_Connection.commit()
#insert FYGDC for MDs into the same table as the FYGDC for the ADs
print "35. Inserting the first year gdc amount into the table tbl_02_FYGDC_By_Staff_Code for the Reporting Offices."
v_Cursor.execute(v_SQLStatementInsertTbl2)
v_Connection.commit()
#create a temporary table with the license weeks by MD
print "36. Creating a Temporary Table with the license weeks for each reporting office between a start and \n" \
      "       end date for sales agents that were licensed on that week."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl10_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl10_A)
    v_Cursor.execute(v_SQLStatementCreateTbl10_A)
    v_Connection.commit()
#Create a temporary table with the FYGDC by MD
print "37. Creating a Temporary Table with the first year gdc by reporting office between a start and an end date."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl10_B)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl10_B)
    v_Cursor.execute(v_SQLStatementCreateTbl10_B)
    v_Connection.commit()
#Create a temporary table with the license weeks and FYGDC by MD
print "38. Creating a Temporary Table with the first year gdc and license weeks between a start and an end date."
try:
    v_Cursor.execute(v_SQLStatementCreateTbl10_C)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl10_C)
    v_Cursor.execute(v_SQLStatementCreateTbl10_C)
    v_Connection.commit()
#Insert the PAPW into tbl_01_PAPW_BY_STAFF_CODE
print "39. Insert the FYGDC and license weeks by reporting offices into tbl_01_PAPW_By_Staff_Code."
v_Cursor.execute(v_SQLStatementInsertTbl1)
v_Connection.commit()
print "40. Creating a Temporary Table with the Reporting Office, employee service number, and first year gdc amt. \n" \
      "       tbl_11_A_TEMP_LOS_1_GT1000_MD"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl11_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl11_A)
    v_Cursor.execute(v_SQLStatementCreateTbl11_A)
    v_Connection.commit()
print "41. Creating a Temporary Table with the Reporting Office, employee service number and the license weeks. \n " \
      "       tbl_11_B_TEMP_LOS_1_GT1000_MD"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl11_B)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl11_B)
    v_Cursor.execute(v_SQLStatementCreateTbl11_B)
    v_Connection.commit()
print "42. Creating a Temporary Table with the Reporting Office, employee service number, license weeks and FYGDC \n"\
      "       tbl_11_C_TEMP_LOS_1_GT1000_MD"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl11_C)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl11_C)
    v_Cursor.execute(v_SQLStatementCreateTbl11_C)
    v_Connection.commit()
print "43. Creating a Temporary Table with the Reporting Office and the Number of LOS 1s that have a PAPW GT 1000 \n" \
      "       tbl_11_D_TEMP_LOS_1_GT1000_MD"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl11_D)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl11_D)
    v_Cursor.execute(v_SQLStatementCreateTbl11_D)
    v_Connection.commit()
print "44. Insert the PAPW GT 1000 Count of LOS 1s into table tbl_04_F_TEMP_LOS_1_GT_1000 with the MD data"
v_Cursor.execute(v_SQLStatementInsertTbl4_F_1)
v_Connection.commit()
print "45. Create a Temporary Table with the AD and MD information for the LOS 1 GT 1000 data. \n" \
      "       tbl_04_G_TEMP_LOS_1_GT_1000"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4_G)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4_G)
    v_Cursor.execute(v_SQLStatementCreateTbl4_G)
    v_Connection.commit()
print "46. Create an instance of MS Excel. "

v_xlApp = win32com.client.Dispatch("Excel.Application")
v_xlApp.AskToUpdateLinks = False
v_xlApp.DisplayAlerts    = False
v_xlApp.ScreenUpdating   = False
v_xlApp.EnableEvents     = False
print "47. Open Excel File "+v_TargetFile
v_xlwb2 = v_xlApp.Workbooks.Open(v_TargetFile)
v_sheet = v_xlwb2.Sheets(1)
print "48. Read in the LOS1 Hiring Target by Month with the service number. "
for v_count_row in range(1,v_sheet.UsedRange.Rows.Count):
    v_sheet.Cells(v_count_row, v_TargetESN).NumberFormat = 'text'
    v_employee_service_number = str(v_sheet.Cells(v_count_row, v_TargetESN).value)
    v_employee_service_number = v_employee_service_number[:-2]
    v_LOS1Hires = str(v_sheet.Cells(v_count_row, v_TargetLOS1Hires).value)
    v_LOS1Hires = v_LOS1Hires[:-2]
    v_RetentionHeadCountByMonth = str(v_sheet.Cells(v_count_row, v_TargetRetentionHeadCountByMonth).value)
    #print v_RetentionHeadCountByMonth[:-2]
    v_TargetLOS1HiresArray.append([str(v_employee_service_number), v_LOS1Hires])
    v_TargetLOS1HiresAndRetention.append([str(v_employee_service_number), v_LOS1Hires, v_RetentionHeadCountByMonth] )
    #print v_TargetLOS1HiresArray

v_xlwb2.Close(True)
v_xlApp.Quit()
print "49. Create a Temporary Table for the AD and MD ESN and the Monthly LOS 1 Hires Target"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4_H)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4_H)
    v_Cursor.execute(v_SQLStatementCreateTbl4_H)
    v_Connection.commit()
print "50. Inserting the data from the Target Spreadsheet to the table \n" \
      "       tbl_04_H_TEMP_LOS_1_GT_1000 "
v_Cursor.executemany(v_SQLStatementInsertTbl4_H_1, v_TargetLOS1HiresArray)
v_Connection.commit()
print "51. Create a Permanent Table with the Agent Information and the Target and Actual for LOS 1 GT 1000"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl4)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl4)
    v_Cursor.execute(v_SQLStatementCreateTbl4)
    v_Connection.commit()
print "52. Creating a Temporary Table of Sales Agent Count with MD information in table : \n"\
      "       tbl_12_A_TEMP_SALES_AGENTS_MD"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl12_A)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl12_A)
    v_Cursor.execute(v_SQLStatementCreateTbl12_A)
    v_Connection.commit()
print "53. Insert MD Sales Agent Counts into table."
v_Cursor.execute(v_SQLStatementInsertTbl5)
v_Connection.commit()
print "54. Create a Permenant Table of the Sales Agent Counts by Staff and Reporting Office \n "\
     "        tbl_05_Sales_Agt_Cnt"
try:
    v_Cursor.execute(v_SQLStatementCreateTbl5)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl5)
    v_Cursor.execute(v_SQLStatementCreateTbl5)
    v_Connection.commit()
print "55. Create a Temporary Table with the LOS Classes and Sales Agents by Staff Code on the End Date \n"\
      "       tbl_13_G_TEMP_Retention_Rate"
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_G)
    v_Connection.commit()
except:
    try:
        time.sleep(60)
        v_Cursor.execute(v_SQLStatementDropTbl13_G)
        v_Cursor.execute(v_SQLStatementCreateTbl13_G)
        v_Connection.commit()
        try:
           time.sleep(60)
           v_Cursor.execute(v_SQLStatementDropTbl13_G)
           v_Cursor.execute(v_SQLStatementCreateTbl13_G)
           v_Connection.commit()
           try:
              time.sleep(60)
              v_Cursor.execute(v_SQLStatementDropTbl13_G)
              v_Cursor.execute(v_SQLStatementCreateTbl13_G)
              v_Connection.commit()
           except:
              time.sleep(60)
              v_Cursor.execute(v_SQLStatementCreateTbl13_G)
              v_Connection.commit()
        except:
           time.sleep(60)
           v_Cursor.execute(v_SQLStatementCreateTbl13_G)
           v_Connection.commit()
    except:
        v_Connection.close()
        sys.exit("THE SCRIPT FAILED.... Please Run Again.")


print "56. Create a Temporary Table with the LOS 3 Sales Agents by Reporting Office on the End Date \n" \
      "       tbl_13_J_TEMP_Retention_Rate"
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_J)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_J)
    v_Cursor.execute(v_SQLStatementCreateTbl13_J)
    v_Connection.commit()
print "57. Create a Temporary Table with the LOS 4 Sales Agents by Reporting Office on the End Date \n" \
      "       tbl_13_K_TEMP_Retention_Rate"
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_K)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_K)
    v_Cursor.execute(v_SQLStatementCreateTbl13_K)
    v_Connection.commit()
print "58. Create a Temporary Table with the LOS 5 Sales Agents by Reporting Office on the End Date \n" \
      "       tbl_13_L_TEMP_Retention_Rate"
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_L)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_L)
    v_Cursor.execute(v_SQLStatementCreateTbl13_L)
    v_Connection.commit()
print "59. Create a Temporary Table with all Reporting Office \n" \
      "       tbl_13_M_TEMP_Retention_Rate"
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_M)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_M)
    v_Cursor.execute(v_SQLStatementCreateTbl13_M)
    v_Connection.commit()
print "60. Create a Temporary Table with all LOS Class Counts by Reporting Office \n" \
      "       tbl_13_N_TEMP_Retention_Rate"
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_N)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_N)
    v_Cursor.execute(v_SQLStatementCreateTbl13_N)
    v_Connection.commit()
print "61. Insert Reporting Office data into the table tbl_13_G_TEMP_Retention_Rate from \n" \
      "       tbl_13_G_TEMP_Retention_Rate "
v_Cursor.execute(v_SQLStatementInsertTbl13_G)
v_Connection.commit()
print "62. Create a Temporary Table will all Employee Service Number, LOS 1 Hire Monthly Target, \n" \
      "       and the LOS 1 Monthly Retention Target  - structure only \n" \
      "       tbl_13_P_TEMP_Retention_Rate "
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_P)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_P)
    v_Cursor.execute(v_SQLStatementCreateTbl13_P)
    v_Connection.commit()
print "63. Insert values into the temporary table \n " \
      "       tbl_13_P_TEMP_Retention_Rate "
v_Cursor.executemany(v_SQLStatementInsertTbl13_P_1, v_TargetLOS1HiresAndRetention)
v_Connection.commit()
print "64. Create a Table will all Retention Data, \n" \
      "       tbl_13_Retention_Rate "
time.sleep(60)
try:
    v_Cursor.execute(v_SQLStatementCreateTbl13_Q)
    v_Connection.commit()
except:
    v_Cursor.execute(v_SQLStatementDropTbl13_Q)
    v_Cursor.execute(v_SQLStatementCreateTbl13_Q)
    v_Connection.commit()
print "Script Finished : "+time.ctime()
v_Connection.close()

