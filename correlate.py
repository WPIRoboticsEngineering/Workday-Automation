import pickle
import odoorpc
from datetime import datetime

logins = pickle.load(open('logins.pickle','rb'))
tlist =   pickle.load( open( "newdat.pickle", "rb" ) ) 

odoo = odoorpc.ODOO('odoo.cs.wpi.edu', port=8069)
odoo.login('wpishop',logins['odoo']['username'],logins['odoo']['password']) 

ai = odoo.env['account.invoice']
po = odoo.env['purchase.order']  
invoice_ids_list = ai.search([])
invoices = ai.browse(invoice_ids_list)

matches = 0;
for i in invoices:
    if i.origin[0:2]=="PO":
        i_po = po.browse(int(i.origin[2:]))
        idate= i_po.date_order.date()
        i_vendor = i_po.partner_id.name
        #print(idate)
        for r in tlist:
            rdate = datetime.strptime(r['date'],'%m/%d/%Y').date()
            #print("\t%s"%(rdate))
            if abs((idate - rdate).days)<15:
                iprice = "%.2f" % (i.amount_total)
                rprice = r['amount']
                if iprice==rprice:
                    matches+=1
                    print("Match!\n%s--%s:%s:%s:%s:%s\n%s\n\n"%(abs((idate - rdate).days),i,i_vendor,i_po,iprice,idate,r))

print(matches)


# Check price and date.
