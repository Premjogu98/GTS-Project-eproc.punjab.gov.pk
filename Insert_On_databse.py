from datetime import datetime
import Global_var
import time
import mysql.connector
import sys , os


def DB_connection():
    mydb = ''
    a = 0
    while a == 0:
        try:
            # File_Location = open("D:\\0 PYTHON EXE SQL CONNECTION & DRIVER PATH\\eproc.punjab.gov.pk\\Location For Database & Driver.txt", "r")
            # TXT_File_AllText = File_Location.read()
            # Local_host = str(TXT_File_AllText).partition("Local_host=")[2].partition(",")[0].strip()
            # Local_user = str(TXT_File_AllText).partition("Local_user=")[2].partition(",")[0].strip()
            # Local_password = str(TXT_File_AllText).partition("Local_password=")[2].partition(",")[0].strip()
            # Local_db = str(TXT_File_AllText).partition("Local_db=")[2].partition(",")[0].strip()
            # Local_charset = str(TXT_File_AllText).partition("Local_charset=")[2].partition("\")")[0].strip()

            mydb = mysql.connector.connect(
                host='185.142.34.92',
                user='ams',
                password='TgdRKAGedt%h',
                db='tenders_db',
                charset='utf8')
            print('SQL Connected Local_connection')
            a = 1
            return mydb
        except mysql.connector.ProgrammingError as e:
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname ,
                  "\n" , exc_tb.tb_lineno)
            a = 0
            time.sleep(10)
            mydb.close()


# def L2L_connection():
#     mydb_L2L = ''
#     a3 = 0
#     while a3 == 0:
#         try:
#             File_Location = open("D:\\0 PYTHON EXE SQL CONNECTION & DRIVER PATH\\eproc.punjab.gov.pk\\Location For Database & Driver.txt" , "r")
#             TXT_File_AllText = File_Location.read()

#             L2L_host = str(TXT_File_AllText).partition("L2L_host=")[2].partition(",")[0].strip()
#             L2L_user = str(TXT_File_AllText).partition("L2L_user=")[2].partition(",")[0].strip()
#             L2L_password = str(TXT_File_AllText).partition("L2L_password=")[2].partition(",")[0].strip()
#             L2L_db = str(TXT_File_AllText).partition("L2L_db=")[2].partition(",")[0].strip()
#             L2L_charset = str(TXT_File_AllText).partition("L2L_charset=")[2].partition("\")")[0].strip()
            
#             mydb_L2L = mysql.connector.connect(
#                 host=str(L2L_host) ,
#                 user=str(L2L_user) ,
#                 passwd=str(L2L_password) ,
#                 database=str(L2L_db) ,
#                 charset=str(L2L_charset))
#             print('SQL Connected L2L_connection')
#             a3 = 1
#             return mydb_L2L
#         except mysql.connector.ProgrammingError as e:
#             exc_type , exc_obj , exc_tb = sys.exc_info()
#             fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
#             print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname ,
#                   "\n" , exc_tb.tb_lineno)
#             a3 = 0
#             time.sleep(10)
#             mydb_L2L.close()
def Error_fun(Error,Function_name,SegFeild):
    mydb = DB_connection()
    mycursor = mydb.cursor()
    sql1 = "INSERT INTO errorlog_tbl(Error_Message,Function_Name,Exe_Name) VALUES('" + str(Error).replace("'","''") + "','" + str(Function_name).replace("'","''")+ "','"+str(SegFeild[31])+"')"
    mycursor.execute(sql1)
    mydb.commit()
    mycursor.close()
    mydb.close()
    return sql1

def check_Duplication(SegFeild):
    global a1
    a1 = 0
    while a1 == 0:
        try:
            mydb = DB_connection()
            mycursor = mydb.cursor()
            if SegFeild[13] != '' and SegFeild[24] != '' and SegFeild[7] != '':
                commandText = "SELECT Posting_Id from asia_tenders_tbl where tender_notice_no = '" + str(SegFeild[13]) + "' and Country = '" + str(SegFeild[7]) + "' and doc_last= '" + str(SegFeild[24]) + "'"
            elif SegFeild[13] != "" and SegFeild[7] != "":
                commandText = "SELECT Posting_Id from asia_tenders_tbl where tender_notice_no = '" + str(SegFeild[13]) + "' and Country = '" + str(SegFeild[7]) + "'"
            elif SegFeild[19] != "" and SegFeild[24] != "" and SegFeild[7] != "":
                commandText = "SELECT Posting_Id from asia_tenders_tbl where short_desc = '" + str(SegFeild[19]) + "' and doc_last = '" + SegFeild[24] + "' and Country = '" + SegFeild[7] + "'"
            else:
                commandText = "SELECT Posting_Id from asia_tenders_tbl where short_desc = '" + str(SegFeild[19]) + "' and Country = '" + str(SegFeild[7]) + "'"
            mycursor.execute(commandText)
            results = mycursor.fetchall()
            mydb.close()
            mycursor.close()
            a1 = 1
            print("Code Reached On check_Duplication")
            return results
        except Exception as e:
            Function_name: str = sys._getframe().f_code.co_name
            Error: str = str(e)
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : ", sys._getframe().f_code.co_name + "--> " + str(e), "\n", exc_type, "\n", fname, "\n",exc_tb.tb_lineno)
            Error_fun(Error,Function_name,SegFeild)
            time.sleep(10)
            a1 = 0


def insert_in_Local(SegFeild):

    results = check_Duplication(SegFeild)
    if len(results) > 0:
        print('Duplicate Tender')
        Global_var.duplicate += 1
        return 1
    else:
        Fileid = create_filename(SegFeild)
    MyLoop = 0
    while MyLoop == 0:
        mydb = DB_connection()
        mycursor = mydb.cursor()
        sql = "INSERT INTO asia_tenders_tbl(Tender_ID,EMail,add1,Country,Maj_Org,tender_notice_no,notice_type,Tenders_details,short_desc,est_cost,currency,doc_cost,doc_last,earnest_money,Financier,tender_doc_file,source)VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        val= (str(Fileid) ,str(SegFeild[1]) , str(SegFeild[2]) , str(SegFeild[7]) , str(SegFeild[12]) , str(SegFeild[13]) , str(SegFeild[14]),
                str(SegFeild[18]) , str(SegFeild[19]) , str(SegFeild[20]) , str(SegFeild[21]) , str(SegFeild[22]), str(SegFeild[24]),str(SegFeild[26]) ,str(SegFeild[27]),
                str(SegFeild[28]) , str(SegFeild[31]))
        try:
            mycursor.execute(sql , val)
            mydb.commit()
            mydb.close()
            mycursor.close()
            Global_var.inserted += 1
            print("Code Reached On insert_in_Local")
            MyLoop = 1
        except Exception as e:
            Function_name :str = sys._getframe().f_code.co_name
            Error : str = str(e)
            Error_fun(Error,Function_name,SegFeild)
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" ,fname , "\n" , exc_tb.tb_lineno)
            MyLoop = 0
            time.sleep(10)
    insert_L2L(SegFeild, Fileid)


def create_filename(SegFeild):
    basename = "PY133"
    Current_dateTime = datetime.now().strftime("%Y%m%d%H%M%S%f")
    Fileid = "".join([basename , Current_dateTime])
     a = 0
     while a == 0:
         try:
             File_path = 'Z:\\' + Fileid + '.html'
             file1 = open(File_path , "w" , encoding='utf-8')
             Html_wala_Tag = "<table align=\"center\" border=\"1\" style=\"width:98%;border-spacing:0;border-collapse: collapse;border:1px solid #666666; margin-top:5px; margin-bottom:5px;\">" +\
                                             "<tr><td colspan=\"2\"; style=\"background-color:#146faf; font-weight: bold; padding:7px;border-bottom:1px solid #666666; color:white;\">Tender Details</td></tr>" +\
                                             "<tr bgcolor=\"#ffffff\" onmouseover=\"this.style.backgroundColor='#def3ff'\" onmouseout=\"this.style.backgroundColor=''\"><td style=\"padding:7px;\">Tender Title</td><td style=\"padding:7px;\">" + str(SegFeild[19]) + "</td></tr>" +\
                                             "<tr bgcolor=\"#ffffff\" onmouseover=\"this.style.backgroundColor='#def3ff'\" onmouseout=\"this.style.backgroundColor=''\"><td style=\"padding:7px;\">Tender Detail</td><td style=\"padding:7px;\">" + str(SegFeild[18])+"</td></tr>" +\
                                             "<tr bgcolor=\"#ffffff\" onmouseover=\"this.style.backgroundColor='#def3ff'\" onmouseout=\"this.style.backgroundColor=''\"><td style=\"padding:7px;\">Department</td><td style=\"padding:7px;\">" + str(SegFeild[12]) + "</td></tr>" +\
                                             "<tr bgcolor=\"#ffffff\" onmouseover=\"this.style.backgroundColor='#def3ff'\" onmouseout=\"this.style.backgroundColor=''\"><td style=\"padding:7px;\">Close Date</td><td style=\"padding:7px;\">" + str(SegFeild[24]) + "</td></tr>" +\
                                             "<tr bgcolor=\"#ffffff\" onmouseover=\"this.style.backgroundColor='#def3ff'\" onmouseout=\"this.style.backgroundColor=''\"><td style=\"padding:7px;\">Tender Notice Attachment </td><td style=\"padding:7px;\">""<a href="+str(SegFeild[4])+" target=\"_blank\">View</a>""</td></tr>" +\
                                            "<tr bgcolor=\"#ffffff\" onmouseover=\"this.style.backgroundColor='#def3ff'\" onmouseout=\"this.style.backgroundColor=''\"><td style=\"padding:7px;\">Bidding Document </td><td style=\"padding:7px;\">""<a href="+str(SegFeild[5])+" target=\"_blank\">View</a>""</td></tr>" + "</tr></table>"
             HTML_File_String = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\"><html xmlns=\"http://www.w3.org/1999/xhtml\">" +\
                                              "<head><link rel=\"shortcut icon\" type=\"image/png\" href=\"https://www.tendersontime.com/favicon.ico\"/></head>" +\
                                              "<body>" + Html_wala_Tag + "</body></html>"
             
             file1.write(str(HTML_File_String))
             file1.close()
             time.sleep(2)
             a = 1
             return Fileid
         except Exception as e:
            Function_name :str = sys._getframe().f_code.co_name
            Error : str = str(e)
            Error_fun(Error,Function_name,SegFeild)
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" ,fname , "\n" , exc_tb.tb_lineno)
            a = 0
            time.sleep(10)


def insert_L2L(SegFeild  , Fileid):
    ncb_icb = "icb"
    added_on = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    search_id = "1"
    cpv_userid = ""
    dms_entrynotice_tblquality_status = '1'
    quality_id = '1'
    quality_addeddate = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    Col1 = 'https://eproc.punjab.gov.pk'
    if SegFeild[7] == "IN":
        Col2 = str(SegFeild[26]) + " * " + str(SegFeild[20])  # For India Only Other Wise Blank
    else:
        Col2 = ''
    Col3 = ''
    Col4 = ''
    Col5 = ''
    file_name = "D:\\Tide\\DocData\\" + Fileid + ".html"
    dms_downloadfiles_tbluser_id = 'DWN5046627'
    # Europe-DWN2554488,India-DWN00541021,Asia-DWN5046627,Africa-DWN302520,North America-DWN1011566,
    # South America-DWN1456891,Semi-Auto-DWN30531073,MFA-DWN0654200
    selector_id = ''
    select_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if SegFeild[36] == "":
        dms_entrynotice_tblstatus = "1"
        dms_downloadfiles_tblsave_status = '1'
        dms_downloadfiles_tblstatus = '1'
        dms_entrynotice_tbl_cqc_status = '1'
    else:
        dms_entrynotice_tblstatus = "2"
        dms_downloadfiles_tblsave_status = '2'
        dms_downloadfiles_tblstatus = '2'
        dms_entrynotice_tbl_cqc_status = '2'
    dms_downloadfiles_tbldatatype = "A"
    dms_entrynotice_tblnotice_type = '2'
    file_id = Fileid
    mydb = DB_connection()
    mycursor = mydb.cursor()
    if SegFeild[12] != "" and SegFeild[19] != "" and SegFeild[24] != "" and SegFeild[7] != "" and SegFeild[2] != "":
        dms_entrynotice_tblcompulsary_qc = "2"
    else:
        dms_entrynotice_tblcompulsary_qc = "1"
        Global_var.QC_Tenders += 1
        sql = "INSERT INTO qctenders_tbl(Source,tender_notice_no,short_desc,doc_last,Maj_Org,Address,doc_path,Country)VALUES(%s,%s,%s,%s,%s,%s,%s,%s) "
        val = (
        str(SegFeild[31]), str(SegFeild[13]), str(SegFeild[19]), str(SegFeild[24]), str(SegFeild[12]), str(SegFeild[2]),
        "http://tottestupload3.s3.amazonaws.com/" + file_id + ".html", str(SegFeild[7]))
        a4 = 0
        while a4 == 0:
            try:
                mydb = DB_connection()
                mycursor = mydb.cursor()
                mycursor.execute(sql , val)
                mydb.commit()
                a4 = 1
                print("Code Reached On QCTenders")
            except Exception as e:
                Function_name: str = sys._getframe().f_code.co_name
                Error: str = e
                Error_fun(Error,Function_name,SegFeild)
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print("Error ON : ", sys._getframe().f_code.co_name + "--> " + str(e), "\n", exc_type, "\n", fname,"\n", exc_tb.tb_lineno)
                a4 = 0
                time.sleep(10)

    sql = "INSERT INTO l2l_tenders_tbl(notice_no,file_id,purchaser_name,deadline,country,description,purchaser_address,purchaser_email,purchaser_url,purchaser_emd,purchaser_value,financier,deadline_two,tender_details,ncbicb,status,added_on,search_id,cpv_value,cpv_userid,quality_status,quality_id,quality_addeddate,source,tender_doc_file,Col1,Col2,Col3,Col4,Col5,file_name,user_id,status_download_id,save_status,selector_id,select_date,datatype,compulsary_qc,notice_type,cqc_status,DocCost,DocLastDate)VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
    val = (str(SegFeild[13]), file_id, str(SegFeild[12]), str(SegFeild[24]), str(SegFeild[7]), str(SegFeild[19]),
           str(SegFeild[2]), str(SegFeild[1]), str(SegFeild[8]), str(SegFeild[26]), str(SegFeild[20]),
           str(SegFeild[27]), str(SegFeild[24]), str(SegFeild[18]), ncb_icb, dms_entrynotice_tblstatus, str(added_on),
           search_id, str(SegFeild[36]), cpv_userid, dms_entrynotice_tblquality_status, quality_id,
           str(quality_addeddate), str(SegFeild[31]), str(SegFeild[28]), Col1, Col2, Col3, Col4, Col5, file_name,
           dms_downloadfiles_tbluser_id, dms_downloadfiles_tblstatus, dms_downloadfiles_tblsave_status, selector_id,
           str(select_date), dms_downloadfiles_tbldatatype, dms_entrynotice_tblcompulsary_qc,
           dms_entrynotice_tblnotice_type, dms_entrynotice_tbl_cqc_status, str(SegFeild[22]), str(SegFeild[41]))
    a5 = 0
    while a5 == 0:
        try:
            mydb = DB_connection()
            mycursor = mydb.cursor()
            mycursor.execute(sql, val)
            mydb.commit()
            mydb.close()
            mycursor.close()
            print("Code Reached On insert_L2L")
            a5 = 1
        except Exception as e:
            Function_name: str = sys._getframe().f_code.co_name
            Error: str = str(e)
            Error_fun(Error,Function_name,SegFeild)
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" ,fname ,"\n" , exc_tb.tb_lineno)
            a5 = 0
            time.sleep(10)