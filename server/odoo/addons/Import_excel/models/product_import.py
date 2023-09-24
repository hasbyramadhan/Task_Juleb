from odoo import models, fields
import base64
import xlrd


class ProductTemplate(models.Model):
    _name = 'product.template'
    _description = 'Product Template'

    name = fields.Char('Name', required=True)
    default_code = fields.Char('Internal Reference')
    barcode = fields.Char('Barcode')
    cost = fields.Float('Cost')
    list_price = fields.Float('Sales Price')
    tracking = fields.Selection([
        ('none', 'No Tracking'),
        ('lot', 'By Lots/Serial Numbers'),
        ('serial', 'By Serial Numbers')],
        string='Tracking')


class ProductImportWizard(models.TransientModel):
    _name = 'product.import.wizard'
    _description = 'Product Import Wizard'

    data_file = fields.Binary(string="Excel File", required=True)
    import_option = fields.Selection([
        ('create', 'Create New Records'),
        ('update', 'Update Existing Records')],
        string='Import Option', required=True, default='create')

    # @api.multi
    def import_product_data(self):
        product_obj = self.env['product.template']

        # Membaca data dari file Excel
        xls_data = base64.b64decode(self.data_file)
        workbook = xlrd.open_workbook(file_contents=xls_data)
        sheet = workbook.sheet_by_index(0)  # Misalnya, menggunakan sheet pertama

        for row in range(1, sheet.nrows):  # Mulai dari baris kedua (baris judul diabaikan)
            row_data = sheet.row_values(row)

            if self.import_option == 'create':
                # Membuat produk baru
                product_obj.create({
                    'name': row_data[0],
                    'default_code': row_data[1],
                    'barcode': row_data[2],
                    'cost': row_data[3],
                    'list_price': row_data[4],
                    'tracking': row_data[5],
                })
            elif self.import_option == 'update':
                # Mencari produk berdasarkan referensi internal (default_code)
                product = product_obj.search([('default_code', '=', row_data[1])])
                if product:
                    # Memperbarui data produk yang sudah ada
                    product.write({
                        'name': row_data[0],
                        'barcode': row_data[2],
                        'cost': row_data[3],
                        'list_price': row_data[4],
                        'tracking': row_data[5],
                    })

        return {'type': 'ir.actions.act_window_close'}

