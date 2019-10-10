from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.remote.webelement import WebElement
from drop_files import drop_files
WebElement.drop_files = drop_files

import time
import pickle
import odoorpc
import tempfile
import base64
from pathlib import Path
from datetime import datetime



class WorkdayInterface:
    def __init__(self,user,password):
        self.driver = webdriver.Chrome()
        self.loginToWorkday(user,password)
        pass
    
    def waitForAndFindM(self,xpath):
        d = self.driver
        WebDriverWait(d,30).until(EC.presence_of_element_located((By.XPATH,xpath)))
        return d.find_elements_by_xpath(xpath)
        
    def waitForAndFindS(self,xpath):
        d = self.driver
        WebDriverWait(d,30).until(EC.presence_of_element_located((By.XPATH,xpath)))
        return d.find_element_by_xpath(xpath)
    
    def findWorkdayIcon(self):
        d = self.driver
        dvs = d.find_elements_by_tag_name('div') 
        for v in dvs:
            if 'workday' in v.value_of_css_property('background-image'):
                return v
        return None
    
    def loginToWorkday(self,user,password):
        print("Logging into workday as '%s'"%(user))
        d = self.driver
        d.get("https://wd5.myworkday.com/wday/authgwy/wpi/login-saml2.htmld")
        #wait for login form to load
        WebDriverWait(d,30).until(EC.presence_of_element_located((By.CLASS_NAME,'btn-primary')))
        self.fill_text_field('i0116',user,True)
        time.sleep(1)
        self.fill_text_field('i0118',password,True)
        time.sleep(1)
        print("\tWaiting for Sign-in Verification by phone..")
        WebDriverWait(d,130).until(EC.presence_of_element_located((By.CLASS_NAME,'workdayHome-ae')))
        
                
        
        return True;
    def navgateToCreateExpenseReport(self):
        print("\tNavigating to expense report page..")
        d=self.driver
        d.get('https://wd5.myworkday.com/wpi/d/home.htmld')
        time.sleep(1)
        WebDriverWait(d,130).until(EC.presence_of_element_located((By.XPATH,".//div[@data-automation-id='landingPageWorkletSelectionOption'][@title='Expenses']")))
        d.find_elements_by_xpath(".//div[@data-automation-id='landingPageWorkletSelectionOption'][@title='Expenses']")[0].click()
        time.sleep(1)
        d.find_element_by_link_text("Create Expense Report").click() 
        #WebDriverWait(d,130).until(EC.presence_of_element_located((By.CLASS_NAME,'WBNG')))
    
    def getListOfPendingExpenses(self):
        print("Getting list of posted transactions..")
        self.navgateToCreateExpenseReport()
        d=self.driver
        time.sleep(1)
        table=self.waitForAndFindS(".//table[@class='mainTable\']")
        rows = table.find_elements(By.TAG_NAME, "tr")
        transactions = []
        for r in rows:
            trans = {}
            cells = r.find_elements(By.TAG_NAME, "td")
            if len(cells)==11:
                trans['date']=cells[3].text
                trans['expense_item']=cells[4].text
                trans['merchant']=cells[5].text
                trans['charge_desc']=cells[6].text
                trans['amount']=cells[7].text
                trans['currency']=cells[8].text
                trans['billing_account']=cells[9].text
                trans['card_last_four']=cells[10].text
                transactions.append(trans)        
        print("\tThere are %s pending expenses"%(len(transactions)))
        return transactions
    
    def createExpenseReportWithRecord(self,recordmap):
        record = recordmap['workday-record']
        print("\tCreating an expense report in Workday for record '%s - %s$'.."%(record['merchant'],record['amount']))
        d=self.driver
        self.navgateToCreateExpenseReport()
        #get table
        time.sleep(1)
        se = self.waitForAndFindS(".//div[@data-automation-id='MainContainer']")
        d.execute_script("return arguments[0].scrollIntoView();", se)
        
        table=self.waitForAndFindS(".//table[@class='mainTable\']")
        rows = table.find_elements(By.TAG_NAME, "tr")
        #iterate through rows untill we match
        #match with date,merchant,price
        for r in rows:
            cells = r.find_elements(By.TAG_NAME, "td")
            if len(cells)==11:
                if cells[3].text == record['date']:
                    if cells[5].text == record['merchant']:
                        if cells[7].text == record['amount']:
                            # We have a match.
                            #Scroll row into view so we can click it
                            d.execute_script("return arguments[0].scrollIntoView();", r)
                            time.sleep(0.5)
                            cells[1].find_element_by_xpath(".//div[@data-automation-id='checkboxPanel']").click()
                            bt = r.find_element_by_xpath("//button[@data-automation-id='wd-CommandButton_uic_okButton']")
                            d.execute_script("return arguments[0].scrollIntoView();", bt)
                            print("\tMatch found! Creating a new report in workday.")
                            bt.click()
                            WebDriverWait(d,30).until(EC.presence_of_element_located((By.XPATH,".//label[contains(text(),'Date')]")))
                            return True
        print("\tCould not find a match in unexpensed items!")
        return False
    

    
    def fill_text_field(self,id,value,ret=False):
        #print("\tFilling text fiend '%s'"%(id))
        d = self.driver
        time.sleep(1)
        try:
            ufield = d.find_elements_by_id(id)[0]
            ufield.clear()
            ufield.send_keys(value)
            time.sleep(1)
            if ret:
                ufield.send_keys(Keys.RETURN)
            return True
        except IndexError:
            print("Could not find Element")
            return False

    def lookupExpenseReportField(self,fieldname):
        d = self.driver
        lab = self.waitForAndFindM(".//label[contains(text(),'%s')]"%(fieldname))
        for l in lab: 
            try:
                field_attr = l.get_attribute('id')
                try:
                    time.sleep(3)
                    return d.find_element_by_xpath("//input[@aria-labelledby='%s']"%(field_attr)) 
                except NoSuchElementException:
                    uid = field_attr.split("-")[0]
                    uid2 = field_attr.split("-")[2]
                    return d.find_element_by_id("%s--%s-input"%(uid,uid2))  
            except NoSuchElementException:
                pass
        return None
    def fillExpenseReportField(self,field_name,contents,ret=True):
        print("\t\tFilling Field '%s' with '%s'"%(field_name,contents))
        field = self.lookupExpenseReportField(field_name)
        if field==None:
            return None
        field.clear()
        field.send_keys(contents)
        time.sleep(1)
        if ret:
            field.send_keys(Keys.RETURN)  

    def navigateToExpenseHeader(self):
        print("\tNavigating to Invoice Header..")
        d = self.driver
        htab = self.waitForAndFindS(".//div[(@data-automation-id='tabLabel') and (text()='Header')]")
        time.sleep(1)
        htab.click()
        time.sleep(4)
        editb = self.waitForAndFindS(".//span[(@title='Edit')]")
        editb.click()
        WebDriverWait(d,30).until(EC.presence_of_element_located((By.XPATH,".//span[(@title='Save')]")))
        return
        
    def saveExpenseReportHeader(self):
        print("\tSaving Expense Report Header..")
        d = self.driver
        savebtn = self.waitForAndFindS(".//span[(@title='Save')]")
        savebtn.click()
        time.sleep(1)
        return
    
    def clickSubmitReport(self):
        print("\tSubmitting Expense Report..")
        d = self.driver
        savebtn = self.waitForAndFindS(".//span[(@title='Submit')]")
        savebtn.click()
        time.sleep(1)
        return
    def attatchFileToReport(self,path):
        print("\tDragging and Dropping Vendor Invoice..")
        d = self.driver
        drop = self.waitForAndFindM(".//div[@data-automation-id='dragDropTarget']")[-1]
        drop.drop_files(str(path))
        return
        
    def submitExpenseReport(self,report):
        try:
            print("Submitting expense report for '%s'"%(ex['odoo-po'].name))
            a_file = oi.downloadAttatchedInvoice(ex)
            if a_file=="":
                print("\tMissing attatchment on %s"%(ex['odoo-po'].name))
                return
            self.createExpenseReportWithRecord(ex)
            time.sleep(3)
            self.fillExpenseReportField('Expense Item','lab supplies')
            time.sleep(1)
            self.fillExpenseReportField('Business Reason','lab supplies')
            time.sleep(1)
            self.fillExpenseReportField('Memo', ex['odoo-po'].name)
            time.sleep(1)
            self.attatchFileToReport(a_file)
            time.sleep(5)
            self.navigateToExpenseHeader()
            time.sleep(3)
            self.fillExpenseReportField('Memo','Lab Supplies')
            time.sleep(1)
            self.fillExpenseReportField('Business Purpose','Supplies')
            time.sleep(1)
            self.saveExpenseReportHeader()
            time.sleep(5)
            self.clickSubmitReport()
            return
        except:
            self.voidCurrentDraftExpenseReport()
            raise
        
    def voidCurrentDraftExpenseReport(self):
        print("\tVoiding Current Expense Report..")
        d = self.driver
        act = self.waitForAndFindS(".//button[(@data-automation-id='relatedActionsButton')]")
        ActionChains(d).move_to_element(act).click().perform()
        time.sleep(3)
        rai = self.waitForAndFindS(".//div[(@data-automation-id='relatedActionsItemLabel') and (@data-automation-label='Expense Report')]")
        ActionChains(d).move_to_element(rai).perform()
        time.sleep(3)
        raii = self.waitForAndFindS(".//div[(@data-automation-id='relatedActionsItemLabel') and (@data-automation-label='Cancel')]")
        ActionChains(d).move_to_element(raii).click().perform()
        time.sleep(3)
        okbtn = self.waitForAndFindS(".//span[(@title='OK')]")
        okbtn.click()
        time.sleep(3)
        return
 
class OdooInterface:
    def __init__(self,username,password):
        print("Connecting to ODOO Database..")
        self.odoo = odoorpc.ODOO('odoo.cs.wpi.edu', port=8069)
        self.odoo.login('wpishop',username,password) 
        self.ai = self.odoo.env['account.invoice']
        self.po = self.odoo.env['purchase.order']  
        
    def correlateRecordsWithOdooInvoices(self,records):
        print("Correlating Pending Expenses with ODOO Invoices..")
        ai = self.ai
        po = self.po
        odoo = self.odoo
        #TODO: we should only serch for invoices newer than the oldest record.
        invoice_ids_list = ai.search([]) 
        invoices = ai.browse(invoice_ids_list)
        # Cull invoices by discarding ones older then oldest posted transaction
        date = datetime.now().date()
        for t in tlist:
            rdate = datetime.strptime(t['date'],'%m/%d/%Y').date()
            if rdate<date:
                date=rdate
        
        
        
        matches = []
        print("\tChecking %s transactions against %s invoices"%(len(records),len(invoices)))
        for i in invoices:

            if i.origin[0:2]=="PO":
                i_po = po.browse(int(i.origin[2:]))
                idate= i_po.date_order.date()
                if ((idate-date).days)>-10:
                    i_vendor = i_po.partner_id.name
                    #print(idate)
                    for r in tlist:
                        rdate = datetime.strptime(r['date'],'%m/%d/%Y').date()
                        #print("\t%s"%(rdate))
                        if abs((idate - rdate).days)<15:
                            iprice = "%.2f" % (i.amount_total)
                            rprice = r['amount']
                            if iprice==rprice:
                                print("\tMatched '%s - %s$' with %s"%(r['merchant'],r['amount'],i.origin))
                                matches.append({'odoo-po':i_po,'odoo-invoice':i,'workday-record':r})
        return matches
        
    def getInvoiceAttatchmentfromInvoiceMessages(self,invoice):
        inv = invoice['odoo-invoice']
        print("\tFinding attatchments in invoice..")
        for m in inv.message_ids: 
            atts = m.attachment_ids
            if len(atts)!=0:
                print("\t\t file '%s' found!"%(atts.datas_fname))
                return {'filename':atts.datas_fname,'data':atts.datas}
        print("\t\tNo Attatchment Found")
        return None
        
    def downloadAttatchedInvoice(self,invoice):
        inv_file = self.getInvoiceAttatchmentfromInvoiceMessages(invoice)
        if inv_file==None:
            return ""
        tempdir = Path(tempfile.gettempdir())
        save_path = tempdir / inv_file['filename']
        print("\tSaving file '%s' to '%s'.."%(inv_file['filename'],tempdir))
        fh = open(save_path, "wb")
        fh.write(base64.b64decode(inv_file['data']))
        fh.close()
        return save_path

logins = pickle.load(open('logins.pickle','rb'))

wi = WorkdayInterface(logins['wpi']['username'],logins['wpi']['password'])
oi = OdooInterface(logins['odoo']['username'],logins['odoo']['password'])

tlist = wi.getListOfPendingExpenses()
corr  = oi.correlateRecordsWithOdooInvoices(tlist)


ex = corr[0]


wi.submitExpenseReport(ex)
#wi.voidCurrentDraftExpenseReport()
    
    
#wi.createExpenseReportWithRecord(tlist[-1])

    
    