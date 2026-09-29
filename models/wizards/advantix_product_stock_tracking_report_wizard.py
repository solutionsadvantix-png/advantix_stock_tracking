import io
import base64
import xlsxwriter
from datetime import datetime, date
from odoo import models, fields, api, _


class StockTrackingReportWizard(models.TransientModel):
    _name = 'stock.tracking.report.wizard'
    _description = 'Stock Tracking Report Wizard'

    from_date = fields.Date(string='From Date', required=True, default=fields.Date.context_today)
    to_date = fields.Date(string='To Date', required=True, default=fields.Date.context_today)

    product_ids = fields.Many2many('product.product', string='Products')
    category_id = fields.Many2one('product.category', string='Product Category')
    location_id = fields.Many2one('stock.location', string='Location', domain="[('usage', '=', 'internal')]")
    lot_id = fields.Many2one('stock.lot', string='Lot/Serial Number')

    def get_opening(self, product_id):
        if not self.from_date:
            return 0.0

        loc_id = self.location_id.id if self.location_id else None

        loc_where = ""
        params = [product_id, self.from_date]

        if loc_id:
            loc_where = " AND (smv.location_id = %s OR smv.location_dest_id = %s) "
            params.extend([loc_id, loc_id])

        if self.lot_id:
            loc_where += " AND sml.lot_id = %s "
            params.append(self.lot_id.id)

        if loc_id:
            case_in = "smv.location_dest_id = %s"
            case_out = "smv.location_id = %s"
            case_params = [loc_id, loc_id]
        else:
            case_in = "dest_loc.usage = 'internal' AND from_loc.usage != 'internal'"
            case_out = "from_loc.usage = 'internal' AND dest_loc.usage != 'internal'"
            case_params = []

        query = f"""
            SELECT SUM(
                CASE 
                    WHEN {case_in} THEN COALESCE(sml.quantity / NULLIF(uom.factor, 0), sml.quantity)
                    WHEN {case_out} THEN -COALESCE(sml.quantity / NULLIF(uom.factor, 0), sml.quantity)
                    ELSE 0
                END
            ) AS opening_balance
            FROM stock_move smv
            JOIN stock_move_line sml ON sml.move_id = smv.id
            LEFT JOIN stock_location from_loc ON smv.location_id = from_loc.id
            LEFT JOIN stock_location dest_loc ON smv.location_dest_id = dest_loc.id
            LEFT JOIN uom_uom uom ON sml.product_uom_id = uom.id
            WHERE sml.product_id = %s
              AND smv.date < %s
              AND smv.state = 'done'
              AND sml.state = 'done'
              {loc_where}
        """

        full_params = case_params + params
        self.env.cr.execute(query, full_params)
        res = self.env.cr.fetchone()

        return res[0] if res and res[0] is not None else 0.0

    def action_generate_xlsx_all_products(self):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        navy_primary = '#102A43'
        slate_accent = '#334E68'
        subtotal_bg = '#D9E2EC'
        zebra_bg = '#F8FAFC'
        border_color = '#BCCCDC'

        prod_hdr_bg = '#BCCCDC'
        prod_hdr_text = '#0F172A'

        doc_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'font_color': navy_primary, 'valign': 'vcenter'
        })

        filter_box_hdr_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': '#FFFFFF',
            'bg_color': slate_accent, 'align': 'center', 'valign': 'vcenter'
        })

        filter_card_lbl_fmt = workbook.add_format({
            'bold': True, 'font_size': 9, 'font_color': navy_primary,
            'bg_color': '#E4E7EB', 'border': 1, 'border_color': border_color,
            'align': 'center', 'valign': 'vcenter'
        })

        filter_card_val_fmt = workbook.add_format({
            'font_size': 9.5, 'font_color': '#102A43',
            'bg_color': '#FFFFFF', 'border': 1, 'border_color': border_color,
            'align': 'center', 'valign': 'vcenter'
        })

        prod_hdr_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_color': prod_hdr_text,
            'bg_color': prod_hdr_bg, 'align': 'left', 'valign': 'vcenter', 'indent': 1,
            'border': 1, 'border_color': border_color
        })

        prod_hdr_open_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': prod_hdr_text,
            'bg_color': prod_hdr_bg, 'align': 'right', 'valign': 'vcenter',
            'num_format': '#,##0.00', 'border': 1, 'border_color': border_color
        })

        tbl_header_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': '#FFFFFF',
            'bg_color': navy_primary, 'border': 1, 'border_color': navy_primary,
            'align': 'center', 'valign': 'vcenter'
        })

        tbl_header_num_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': '#FFFFFF',
            'bg_color': navy_primary, 'border': 1, 'border_color': navy_primary,
            'align': 'right', 'valign': 'vcenter'
        })

        line_even_fmt = workbook.add_format({
            'border': 1, 'border_color': border_color, 'align': 'center', 'valign': 'vcenter', 'font_size': 9.5
        })

        line_odd_fmt = workbook.add_format({
            'border': 1, 'border_color': border_color, 'bg_color': zebra_bg,
            'align': 'center', 'valign': 'vcenter', 'font_size': 9.5
        })

        num_even_fmt = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'border_color': border_color,
            'align': 'right', 'valign': 'vcenter', 'font_size': 9.5
        })

        num_odd_fmt = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'border_color': border_color,
            'bg_color': zebra_bg, 'align': 'right', 'valign': 'vcenter', 'font_size': 9.5
        })

        prod_subtotal_lbl = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': navy_primary,
            'bg_color': subtotal_bg, 'border': 1, 'border_color': border_color,
            'align': 'right', 'valign': 'vcenter'
        })

        prod_subtotal_num = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': navy_primary,
            'num_format': '#,##0.00', 'bg_color': subtotal_bg,
            'border': 1, 'border_color': border_color, 'align': 'right', 'valign': 'vcenter'
        })

        sheet = workbook.add_worksheet('Grouped Stock Tracking')
        sheet.hide_gridlines(2)

        sheet.set_column('A:A', 18)  # Date
        sheet.set_column('B:B', 18)  # Transfer
        sheet.set_column('C:C', 18)  # Transfer Type
        sheet.set_column('D:D', 20)  # Operation Type
        sheet.set_column('E:E', 22)  # From
        sheet.set_column('F:F', 22)  # To
        sheet.set_column('G:G', 14)  # In
        sheet.set_column('H:H', 14)  # Out
        sheet.set_column('I:I', 16)  # Balance
        sheet.set_column('J:J', 22)  # Partner
        sheet.set_column('K:K', 18)  # Lot

        sheet.set_row(1, 26)
        sheet.write(1, 0, 'STOCK TRACKING REPORT (GROUPED BY PRODUCT)', doc_title_fmt)

        sheet.merge_range('A3:D3', 'REPORT APPLIED FILTERS', filter_box_hdr_fmt)

        sheet.write(3, 0, 'From Date', filter_card_lbl_fmt)
        sheet.write(4, 0, str(self.from_date or 'All Dates'), filter_card_val_fmt)

        sheet.write(3, 1, 'To Date', filter_card_lbl_fmt)
        sheet.write(4, 1, str(self.to_date or 'All Dates'), filter_card_val_fmt)

        sheet.write(3, 2, 'Location', filter_card_lbl_fmt)
        sheet.write(4, 2, self.location_id.display_name if self.location_id else 'All Locations', filter_card_val_fmt)

        sheet.write(3, 3, 'Category / Selection', filter_card_lbl_fmt)
        cat_or_prod = self.category_id.name if self.category_id else (
            'Selected Products' if self.product_ids else 'All Products')
        sheet.write(4, 3, cat_or_prod, filter_card_val_fmt)

        if self.product_ids:
            products = self.product_ids
        elif self.category_id:
            products = self.env['product.product'].search([
                ('categ_id', 'child_of', self.category_id.id),
                ('type', '=', 'consu')
            ])
        else:
            products = self.env['product.product'].search([('type', '=', 'consu')])

        current_row = 6
        subtotal_rows = []

        headers = [
            'Date', 'Transfer', 'Transfer Type', 'Operation Type',
            'From', 'To', 'In', 'Out', 'Balance', 'Partner', 'Lot'
        ]

        for product in products:
            data = self._get_product_report_data(product)
            report_lines = data.get('report_data', [])

            opening_bal = self.get_opening(product.id)
            if not report_lines and opening_bal == 0.0:
                continue

            product_display = f"PRODUCT: [{product.default_code}] {product.name}" if product.default_code else f"PRODUCT: {product.name}"

            sheet.set_row(current_row, 24)
            sheet.merge_range(current_row, 0, current_row, 5, product_display, prod_hdr_title_fmt)
            sheet.write(current_row, 6, 'Opening Bal:', prod_hdr_title_fmt)
            sheet.write(current_row, 7, opening_bal, prod_hdr_open_fmt)
            sheet.merge_range(current_row, 8, current_row, 10, '', prod_hdr_title_fmt)
            current_row += 1

            sheet.set_row(current_row, 22)
            for col_idx, text in enumerate(headers):
                fmt = tbl_header_num_fmt if text in ['In', 'Out', 'Balance'] else tbl_header_fmt
                sheet.write(current_row, col_idx, text, fmt)

            current_row += 1
            prod_start_row = current_row + 1  # 1-indexed for Excel SUM
            running_balance = opening_bal

            for idx, line in enumerate(report_lines):
                sheet.set_row(current_row, 20)
                f_line = line_odd_fmt if idx % 2 == 0 else line_even_fmt
                f_num = num_odd_fmt if idx % 2 == 0 else num_even_fmt

                in_qty = line.get('in_qty', 0.0)
                out_qty = line.get('out_qty', 0.0)
                running_balance += (in_qty - out_qty)

                sheet.write(current_row, 0, line.get('date', ''), f_line)
                sheet.write(current_row, 1, line.get('transfer', ''), f_line)
                sheet.write(current_row, 2, line.get('transfer_type', ''), f_line)
                sheet.write(current_row, 3, line.get('operation_type', ''), f_line)
                sheet.write(current_row, 4, line.get('source_loc', ''), f_line)
                sheet.write(current_row, 5, line.get('destination_loc', ''), f_line)
                sheet.write(current_row, 6, in_qty, f_num)
                sheet.write(current_row, 7, out_qty, f_num)
                sheet.write(current_row, 8, running_balance, f_num)
                sheet.write(current_row, 9, line.get('partner', ''), f_line)
                sheet.write(current_row, 10, line.get('lot', ''), f_line)

                current_row += 1

            prod_end_row = current_row
            sheet.set_row(current_row, 22)
            sheet.merge_range(current_row, 0, current_row, 5, f'Total for {product.name}:', prod_subtotal_lbl)

            if prod_end_row >= prod_start_row:
                sheet.write_formula(current_row, 6, f"=SUM(G{prod_start_row}:G{prod_end_row})", prod_subtotal_num)
                sheet.write_formula(current_row, 7, f"=SUM(H{prod_start_row}:H{prod_end_row})", prod_subtotal_num)
            else:
                sheet.write(current_row, 6, 0.0, prod_subtotal_num)
                sheet.write(current_row, 7, 0.0, prod_subtotal_num)

            sheet.write(current_row, 8, running_balance, prod_subtotal_num)
            sheet.write(current_row, 9, '', prod_subtotal_lbl)
            sheet.write(current_row, 10, '', prod_subtotal_lbl)

            subtotal_rows.append(current_row + 1)
            current_row += 2

        workbook.close()
        output.seek(0)

        file_data = base64.b64encode(output.read())
        attachment = self.env['ir.attachment'].create({
            'name': f'Master_Stock_Tracking_{fields.Date.today()}.xlsx',
            'type': 'binary',
            'datas': file_data,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }

    def _get_product_report_data(self, product):
        domain_base = [('product_id', '=', product.id), ('state', '=', 'done')]

        if self.location_id:
            domain_base += [
                '|',
                ('location_id', 'child_of', self.location_id.id),
                ('location_dest_id', 'child_of', self.location_id.id)
            ]

        if self.lot_id:
            domain_base.append(('lot_ids', 'in', [self.lot_id.id]))

        opening_domain = domain_base + [('date', '<', self.from_date)]
        opening_moves = self.env['stock.move.line'].search(opening_domain)
        opening_qty = 0.0

        for move in opening_moves:
            qty = move.quantity
            if self.location_id:
                if move.location_dest_id.id == self.location_id.id:
                    opening_qty += qty
                if move.location_id.id == self.location_id.id:
                    opening_qty -= qty
            else:
                if move.location_dest_id.usage == 'internal':
                    opening_qty += qty
                if move.location_id.usage == 'internal':
                    opening_qty -= qty

        period_domain = domain_base + [
            ('date', '>=', self.from_date),
            ('date', '<=', self.to_date)
        ]
        period_moves = self.env['stock.move.line'].search(period_domain, order='date asc')

        report_lines = []
        running_balance = opening_qty

        for move in period_moves:
            qty = move.quantity
            in_qty = 0.0
            out_qty = 0.0

            if self.location_id:
                if move.location_dest_id.id == self.location_id.id:
                    in_qty = qty
                if move.location_id.id == self.location_id.id:
                    out_qty = qty
            else:
                if move.location_dest_id.usage == 'internal' and move.location_id.usage != 'internal':
                    in_qty = qty
                elif move.location_id.usage == 'internal' and move.location_dest_id.usage != 'internal':
                    out_qty = qty
                else:
                    in_qty = qty
                    out_qty = qty

            running_balance += (in_qty - out_qty)

            report_lines.append({
                'date': move.date.strftime('%Y-%m-%d %H:%M:%S') if move.date else '',
                'transfer': move.picking_id.name or move.reference or '',
                'operation_type': move.picking_id.picking_type_id.name or '',
                'transfer_type': move.picking_id.picking_type_id.code or '',
                'source_loc': move.location_id.display_name,
                'destination_loc': move.location_dest_id.display_name,
                'in_qty': in_qty,
                'out_qty': out_qty,
                'balance': running_balance,
                'partner': move.picking_id.partner_id.name or '',
                'lot': ', '.join(move.lot_id.mapped('name')) if move.lot_id else '',
            })

        return {
            'product': product.display_name,
            'location': self.location_id.display_name if self.location_id else 'All Locations',
            'from_date': str(self.from_date or ''),
            'to_date': str(self.to_date or ''),
            'opening': opening_qty,
            'closing': running_balance,
            'report_data': report_lines,
        }