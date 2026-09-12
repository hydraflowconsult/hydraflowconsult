import os
import json
import sqlite3
from datetime import datetime, timedelta
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

class InvoiceGeneratorPro:
    """Professional PDF Invoice Generator - HydraFlow Consult"""
    
    # Template Definitions
    TEMPLATES = {
        'hydraflow': {
            'name': 'HYDRAFLOW CONSULT GHANA',
            'tagline': 'Water • Environment • Sustainability • Engineering',
            'address': '5th Avenue, McCarthy Hill',
            'city': 'Accra, Ghana',
            'phone': '0544139637',
            'email': 'abydove1@gmail.com',
            'tax_id': 'C1234567890',
            'ceo': 'Abigail Twenewaa, M.Sc.',
            'credentials': 'M.Sc. Water Supply & Environmental Sanitation',
            'certifications': 'GSA Technical Committee Member | KNUST Researcher',
            'primary_color': '#0066cc',
            'secondary_color': '#2e7d32',
            'logo_url': None
        }
    }
    
    def __init__(self, template='hydraflow', db_path='invoices.db', email_config=None):
        """Initialize InvoiceGeneratorPro"""
        self.template = self.TEMPLATES.get(template, template)
        self.db_path = db_path
        self.email_config = email_config or {}
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        self._init_database()
    
    def _setup_custom_styles(self):
        """Define all custom paragraph styles"""
        primary = colors.HexColor(self.template['primary_color'])
        secondary = colors.HexColor(self.template['secondary_color'])
        dark_neutral = colors.HexColor("#1a252f")
        light_bg = colors.HexColor("#f8f9fa")
        
        self.meta_style = ParagraphStyle(
            'MetaText',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            textColor=dark_neutral
        )
        
        self.title_style = ParagraphStyle(
            'CompanyTitle',
            parent=self.styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=16,
            leading=20,
            textColor=primary
        )
        
        self.section_heading = ParagraphStyle(
            'SectionHeading',
            parent=self.styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=12,
            leading=16,
            textColor=secondary,
            spaceBefore=10,
            spaceAfter=10
        )
        
        self.table_header_style = ParagraphStyle(
            'TableHeader',
            parent=self.styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=10,
            leading=12,
            textColor=colors.white
        )
        
        self.table_cell_style = ParagraphStyle(
            'TableCell',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            textColor=dark_neutral
        )
        
        self.table_cell_right = ParagraphStyle(
            'TableCellRight',
            parent=self.table_cell_style,
            alignment=2
        )
        
        self.table_header_right = ParagraphStyle(
            'TableHeaderRight',
            parent=self.table_header_style,
            alignment=2
        )
        
        self.credentials_style = ParagraphStyle(
            'Credentials',
            parent=self.meta_style,
            fontSize=9,
            textColor=secondary,
            fontName='Helvetica-Oblique'
        )
        
        self.primary_color = primary
        self.secondary_color = secondary
        self.light_bg = light_bg
    
    def _init_database(self):
        """Initialize SQLite database for invoice tracking"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # Invoices table
        c.execute('''CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY,
            invoice_number TEXT UNIQUE,
            client_name TEXT,
            client_email TEXT,
            invoice_date TEXT,
            due_date TEXT,
            subtotal REAL,
            tax_amount REAL,
            discount_amount REAL,
            total REAL,
            status TEXT DEFAULT 'draft',
            payment_date TEXT,
            payment_method TEXT,
            created_at TEXT,
            updated_at TEXT
        )''')
        
        # Invoice items table
        c.execute('''CREATE TABLE IF NOT EXISTS invoice_items (
            id INTEGER PRIMARY KEY,
            invoice_number TEXT,
            description TEXT,
            quantity REAL,
            unit_price REAL,
            amount REAL,
            FOREIGN KEY(invoice_number) REFERENCES invoices(invoice_number)
        )''')
        
        # Payment tracking table
        c.execute('''CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY,
            invoice_number TEXT,
            payment_date TEXT,
            amount REAL,
            method TEXT,
            reference TEXT,
            notes TEXT,
            FOREIGN KEY(invoice_number) REFERENCES invoices(invoice_number)
        )''')
        
        conn.commit()
        conn.close()
    
    def _create_header(self, invoice_num, invoice_date, due_date=None):
        """Create company and invoice header section"""
        if due_date is None:
            due_date = "Upon Receipt"
        
        company_info = f'''<b>{self.template['name']}</b><br/>
<i>{self.template['credentials']}</i><br/>
{self.template['tagline']}<br/>
{self.template['address']}<br/>
{self.template['city']}<br/>
Tel: {self.template['phone']}<br/>
Email: {self.template['email']}<br/>
Tax ID: {self.template['tax_id']}<br/>
<font size=8><i>{self.template['certifications']}</i></font>'''
        
        invoice_meta = f'''<b>INVOICE</b><br/><br/>
<b>Invoice #:</b> {invoice_num}<br/>
<b>Date:</b> {invoice_date}<br/>
<b>Due Date:</b> {due_date}<br/>
<b>Currency:</b> GHS (GH₵)'''
        
        header_data = [
            [
                Paragraph(company_info, self.meta_style),
                Paragraph(invoice_meta, ParagraphStyle('RightMeta', parent=self.meta_style, alignment=2))
            ]
        ]
        
        header_table = Table(header_data, colWidths=[300, 200])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('PADDING', (0, 0), (-1, -1), 0),
        ]))
        
        return header_table
    
    def _create_bill_to(self, client_name, client_org, client_email=None, client_phone=None):
        """Create 'Bill To' section"""
        bill_to = f'<b>Client Name:</b> {client_name}<br/><b>Company/Organization:</b> {client_org}'
        if client_email:
            bill_to += f'<br/><b>Email:</b> {client_email}'
        if client_phone:
            bill_to += f'<br/><b>Phone:</b> {client_phone}'
        return bill_to
    
    def _create_items_table(self, items, tax_rate=0, discount_amount=0):
        """Create itemized services table with tax and discount"""
        table_data = [[
            Paragraph("<b>Service Description</b>", self.table_header_style),
            Paragraph("<b>Qty</b>", self.table_header_right),
            Paragraph("<b>Unit Price</b>", self.table_header_right),
            Paragraph("<b>Amount</b>", self.table_header_right)
        ]]
        
        subtotal = 0
        for item in items:
            if len(item) == 2:
                desc, amt = item
                qty, unit_price = 1, amt
            else:
                desc, qty, unit_price = item
                amt = qty * unit_price
            
            subtotal += amt
            table_data.append([
                Paragraph(desc, self.table_cell_style),
                Paragraph(f"{qty}", self.table_cell_right),
                Paragraph(f"GH₵ {unit_price:,.2f}", self.table_cell_right),
                Paragraph(f"GH₵ {amt:,.2f}", self.table_cell_right)
            ])
        
        # Subtotal row
        table_data.append([
            Paragraph("", self.table_cell_style),
            Paragraph("", self.table_cell_right),
            Paragraph("<b>Subtotal:</b>", self.table_cell_right),
            Paragraph(f"<b>GH₵ {subtotal:,.2f}</b>", self.table_cell_right)
        ])
        
        # Discount row
        if discount_amount > 0:
            table_data.append([
                Paragraph("", self.table_cell_style),
                Paragraph("", self.table_cell_right),
                Paragraph("<b>Discount:</b>", self.table_cell_right),
                Paragraph(f"<b>-GH₵ {discount_amount:,.2f}</b>", self.table_cell_right)
            ])
            subtotal -= discount_amount
        
        # Tax row
        tax_amount = subtotal * (tax_rate / 100) if tax_rate > 0 else 0
        if tax_rate > 0:
            table_data.append([
                Paragraph("", self.table_cell_style),
                Paragraph("", self.table_cell_right),
                Paragraph(f"<b>Tax ({tax_rate}%):</b>", self.table_cell_right),
                Paragraph(f"<b>GH₵ {tax_amount:,.2f}</b>", self.table_cell_right)
            ])
        
        # Total row
        total = subtotal + tax_amount
        table_data.append([
            Paragraph("", self.table_cell_style),
            Paragraph("", self.table_cell_right),
            Paragraph("<b>Total Balance Due:</b>", ParagraphStyle('TotalLabel', parent=self.table_cell_right, fontName='Helvetica-Bold')),
            Paragraph(f"<b>GH₵ {total:,.2f}</b>", ParagraphStyle('TotalVal', parent=self.table_cell_right, fontName='Helvetica-Bold', fontSize=11))
        ])
        
        item_table = Table(table_data, colWidths=[280, 60, 110, 110])
        
        # Style the table
        t_style = [
            ('BACKGROUND', (0, 0), (-1, 0), self.primary_color),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            ('GRID', (0, 0), (-1, -2), 0.5, colors.HexColor("#e0e0e0")),
            ('LINEABOVE', (0, -2), (-1, -2), 1, colors.HexColor("#e0e0e0")),
            ('LINEBELOW', (0, -1), (-1, -1), 1.5, self.primary_color),
            ('TOPPADDING', (0, -1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, -1), (-1, -1), 10),
            ('BACKGROUND', (0, -1), (-1, -1), self.light_bg),
        ]
        
        for idx in range(1, len(items) + 1):
            if idx % 2 == 0:
                t_style.append(('BACKGROUND', (0, idx), (-1, idx), self.light_bg))
        
        item_table.setStyle(TableStyle(t_style))
        return item_table, {'subtotal': subtotal - tax_amount, 'tax': tax_amount, 'discount': discount_amount, 'total': total}
    
    def _save_to_database(self, invoice_num, client_name, client_email, invoice_date, 
                         due_date, items, tax_rate, discount_amount, total):
        """Save invoice to database"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        subtotal = sum(item[2] * item[1] if len(item) == 3 else item[1] for item in items)
        tax_amount = subtotal * (tax_rate / 100)
        
        try:
            c.execute('''INSERT INTO invoices 
                        (invoice_number, client_name, client_email, invoice_date, due_date, 
                         subtotal, tax_amount, discount_amount, total, status, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                     (invoice_num, client_name, client_email, invoice_date, due_date,
                      subtotal, tax_amount, discount_amount, total, 'draft',
                      datetime.now().isoformat(), datetime.now().isoformat()))
            
            for item in items:
                if len(item) == 2:
                    desc, amt = item
                    c.execute('''INSERT INTO invoice_items 
                                (invoice_number, description, quantity, unit_price, amount)
                                VALUES (?, ?, ?, ?, ?)''',
                             (invoice_num, desc, 1, amt, amt))
                else:
                    desc, qty, unit_price = item
                    c.execute('''INSERT INTO invoice_items 
                                (invoice_number, description, quantity, unit_price, amount)
                                VALUES (?, ?, ?, ?, ?)''',
                             (invoice_num, desc, qty, unit_price, qty * unit_price))
            
            conn.commit()
        except sqlite3.IntegrityError:
            print(f"Invoice {invoice_num} already exists in database")
        finally:
            conn.close()
    
    def _send_email(self, recipient_email, invoice_num, pdf_path):
        """Send invoice via email"""
        if not self.email_config:
            print("⚠️  Email config not set. Skipping email send.")
            return False
        
        try:
            msg = MIMEMultipart()
            msg['From'] = self.email_config.get('username')
            msg['To'] = recipient_email
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = f'Invoice {invoice_num} - HydraFlow Consult Ghana'
            
            body = f'''Dear Valued Client,

Please find attached your professional invoice {invoice_num} from HydraFlow Consult Ghana.

Invoice Details:
- Invoice Number: {invoice_num}
- Invoice Date: {datetime.now().strftime('%B %d, %Y')}

Thank you for partnering with us towards sustainable water solutions!

Best regards,
Abigail Twenewaa, M.Sc.
CEO & Founder
HydraFlow Consult Ghana
{self.template['email']}
{self.template['phone']}
Accra, Ghana'''
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach PDF
            with open(pdf_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(pdf_path)}')
                msg.attach(part)
            
            # Send email
            server = smtplib.SMTP(self.email_config.get('smtp_server'), self.email_config.get('smtp_port'))
            server.starttls()
            server.login(self.email_config.get('username'), self.email_config.get('password'))
            server.send_message(msg)
            server.quit()
            
            print(f"✓ Email sent to {recipient_email}")
            return True
        except Exception as e:
            print(f"❌ Email failed: {e}")
            return False
    
    def create_invoice(self, filename, invoice_num, invoice_date, client_name, client_org,
                      items, client_email=None, client_phone=None, due_date=None, notes=None,
                      tax_rate=0, discount_amount=0, send_email=False):
        """Generate professional PDF invoice"""
        os.makedirs(os.path.dirname(filename) if os.path.dirname(filename) else '.', exist_ok=True)
        
        doc = SimpleDocTemplate(filename, pagesize=letter, title=f"Invoice {invoice_num}")
        story = []
        
        # Add header
        story.append(self._create_header(invoice_num, invoice_date, due_date))
        story.append(Spacer(1, 25))
        
        # Add bill to
        story.append(Paragraph("<b>BILL TO:</b>", self.section_heading))
        bill_to_text = self._create_bill_to(client_name, client_org, client_email, client_phone)
        story.append(Paragraph(bill_to_text, self.meta_style))
        story.append(Spacer(1, 20))
        
        # Add items table
        item_table, amounts = self._create_items_table(items, tax_rate, discount_amount)
        story.append(item_table)
        story.append(Spacer(1, 30))
        
        # Add payment terms
        terms_text = f'''<b>Payment Instructions & Terms:</b><br/>
1. Make payment to <b>HydraFlow Consult Ghana</b><br/>
2. Reference Invoice #{invoice_num}<br/>
3. Contact {self.template['phone']} for mobile money options<br/>
4. Payment due within 30 days unless otherwise specified'''
        story.append(Paragraph(terms_text, self.meta_style))
        story.append(Spacer(1, 20))
        
        if notes:
            story.append(Paragraph("<b>Additional Notes:</b>", self.section_heading))
            story.append(Paragraph(notes, self.meta_style))
            story.append(Spacer(1, 20))
        
        # Add thank you
        thank_you = Paragraph(
            f"<i>Thank you for partnering with HydraFlow Consult for sustainable water solutions!</i>",
            ParagraphStyle('ThankYou', parent=self.meta_style, alignment=1, textColor=self.secondary_color)
        )
        story.append(thank_you)
        
        doc.build(story)
        print(f"✓ Invoice generated: {filename}")
        
        # Save to database
        self._save_to_database(invoice_num, client_name, client_email, invoice_date, 
                              due_date or "Upon Receipt", items, tax_rate, discount_amount, amounts['total'])
        
        # Send email if requested
        if send_email and client_email:
            self._send_email(client_email, invoice_num, filename)
        
        return amounts


# Analytics & Reporting Functions
class InvoiceAnalytics:
    """Generate reports and analytics from invoice data"""
    
    def __init__(self, db_path='invoices.db'):
        self.db_path = db_path
    
    def get_monthly_revenue(self, year, month):
        """Get total revenue for a specific month"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('''SELECT SUM(total) FROM invoices 
                    WHERE strftime('%Y-%m', invoice_date) = ?''',
                 (f"{year:04d}-{month:02d}",))
        
        result = c.fetchone()
        conn.close()
        return result[0] or 0
    
    def get_client_spending(self, client_name):
        """Get total spending by a specific client"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('SELECT SUM(total) FROM invoices WHERE client_name = ?', (client_name,))
        result = c.fetchone()
        conn.close()
        return result[0] or 0
    
    def get_invoice_status_summary(self):
        """Get summary of invoice statuses"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('SELECT status, COUNT(*), SUM(total) FROM invoices GROUP BY status')
        results = c.fetchall()
        conn.close()
        
        return {row[0]: {'count': row[1], 'total': row[2]} for row in results}
    
    def get_top_clients(self, limit=10):
        """Get top clients by revenue"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('''SELECT client_name, COUNT(*) as invoice_count, SUM(total) as total_revenue
                    FROM invoices GROUP BY client_name ORDER BY total_revenue DESC LIMIT ?''',
                 (limit,))
        
        results = c.fetchall()
        conn.close()
        return results
    
    def generate_monthly_report(self, year, month):
        """Generate comprehensive monthly report"""
        revenue = self.get_monthly_revenue(year, month)
        status_summary = self.get_invoice_status_summary()
        top_clients = self.get_top_clients(5)
        
        report = {
            'month': f"{year:04d}-{month:02d}",
            'total_revenue': revenue,
            'invoice_statuses': status_summary,
            'top_clients': [{'name': row[0], 'invoices': row[1], 'revenue': row[2]} for row in top_clients],
            'generated_at': datetime.now().isoformat()
        }
        
        return report


if __name__ == '__main__':
    # Initialize generator
    generator = InvoiceGeneratorPro(template='hydraflow')
    
    # Demo items
    demo_items = [
        ("Phase 1: Hydrogeological Technical Feasibility Study & Site Assessment", 1, 4500.00),
        ("Water Quality Testing & Comprehensive Laboratory Assessment", 2, 600.00),
        ("Environmental Sustainability Advisory & Project Planning Report", 1, 2800.00)
    ]
    
    os.makedirs('generated', exist_ok=True)
    
    # Create invoice with tax and discount
    amounts = generator.create_invoice(
        'generated/hydraflow_invoice_pro.pdf',
        'HF-2026-001',
        'September 11, 2026',
        'Emmanuel Mensah',
        'Accra Infrastructure Dev',
        demo_items,
        client_email='emmanuel@accradev.com',
        client_phone='0501234567',
        due_date='October 11, 2026',
        notes='Professional services for sustainable water solutions',
        tax_rate=15,
        discount_amount=500,
        send_email=False
    )
    
    print(f"\n💰 Invoice Summary:")
    print(f"   Subtotal: GH₵ {amounts['subtotal']:,.2f}")
    print(f"   Tax (15%): GH₵ {amounts['tax']:,.2f}")
    print(f"   Discount: GH₵ {amounts['discount']:,.2f}")
    print(f"   Total: GH₵ {amounts['total']:,.2f}")
