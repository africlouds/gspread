import gspread
from oauth2client.service_account import ServiceAccountCredentials
import psycopg2
import psycopg2.extras
import datetime
import sys, traceback


scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('ChillingtonAccounts-874e39d19c13.json', scope)

gc = gspread.authorize(credentials)

conn = psycopg2.connect(dbname="chillington_db", user="chillington", host="localhost", port="9999", password="chillington")
cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)


wks = gc.open("Test").worksheet("Test")
def update_names(wks):
	for i, row in enumerate(values):
		try:
			product = get_product(int(row[0]))
			print product['name_template']
			wks.update_acell('B%i' % (i+1), product['name_template'])
		except:
			pass

def get_product(openerp_id):
	product = None
	cur.execute("SELECT * FROM product_product WHERE id=%i" % openerp_id)
	for product in cur:
		pass
	return product

def get_prod_balance(product, fro, to, direction):
	query = " SELECT SUM(product_uos_qty) FROM stock_move WHERE product_id=%i AND date > '%s' AND date <= '%s' AND location_dest_id=%i" % (product, fro, to, 12 if direction=="in" else 11)
	cur.execute(query)
	res = cur.fetchone()	
	return res['sum'] if res['sum'] else 0
#read the whole Foundry sheet

def get_last_column(wks,row):
	row = wks.row_values(row)
	for i, val in reversed(list(enumerate(row))):
		if val != "":
			return i + 1

def get_latest_date(row, col):
	return datetime.datetime.strptime(wks.cell(row, col).value, "%d/%m/%Y")

def get_next_date(current_date, days):
	return current_date + datetime.timedelta(days=days)
	
def update_balances(wks):
	last_col = get_last_column(wks, 1)
	current_date = get_latest_date(1, last_col)
	next_date = get_next_date(current_date, 7)
	rows = wks.get_all_values()
	wks.update_cell(1, last_col + 1, next_date.strftime('%d/%m/%Y'))
	for i, row in enumerate(rows):
		try:
			openerp_id = int(row[0])
			out_bal = get_prod_balance(openerp_id, current_date, next_date, "out")
			in_bal = get_prod_balance(openerp_id, current_date, next_date, "in")
			wks.update_cell(i+1, last_col + 1, "In: %i,Out: %i" % (out_bal, in_bal))
		except:
			traceback.print_exc(file=sys.stdout)


values = wks.get_all_values()
#update_names(wks)
#wks.update_cell(i+1, len(row), row[-1])
last_col = get_last_column(wks, 1)
print last_col
current_date = get_latest_date(1, last_col)
next_date = get_next_date(current_date, 7)
print current_date
print next_date
print get_prod_balance(195, current_date, next_date, "out")

for i in range(40):
	update_balances(wks)



