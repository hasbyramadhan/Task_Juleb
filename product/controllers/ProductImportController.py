from odoo import http
from odoo.http import request

class ProductTemplateImportController(http.Controller):
    @http.route('/import_product_template', type='http', auth='user')
    def import_product_template_page(self, **kw):
        return request.render('product.view_product_template_import_form', {})

    @http.route('/import_product_template', type='http', auth='user', methods=['POST'], website=True)
    def import_product_template(self, **kw):
        excel_file = kw.get('excel_file')

        if excel_file:
            # Membuat wizard dan mengisi field excel_file
            wizard = request.env['product.template.import'].create({'excel_file': excel_file})
            # Melakukan impor
            wizard.import_product_template()

        return request.redirect('/import_product_template')
