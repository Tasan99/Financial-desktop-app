import pandas as pd
from pathlib import Path
import warnings
import calendar

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


class FinancialDataEngine:
    def __init__(self, file_path):
        """Veriyi yükler, temizler ve motoru hazırlayarak sabitleri tanımlar."""
        self.load_and_clean_data(file_path)
        self.define_constants()

    def load_and_clean_data(self, file_path):
        df_wide = pd.read_excel(file_path, sheet_name='Raw Data', header=5)

        if 'Unnamed: 0' in df_wide.columns:
            df_wide = df_wide.drop(columns=['Unnamed: 0'])

        self.df = df_wide.melt(
            id_vars=['Branch', 'Region', 'Account'],
            var_name='Period',
            value_name='Value'
        )

        self.df['Month_Num'] = self.df['Period'].str.replace('M', '', regex=False).astype(int)
        self.df['Year'] = self.df['Month_Num'].apply(lambda x: 2023 if x <= 12 else 2024)
        self.df['Month'] = self.df['Month_Num'].apply(lambda x: x if x <= 12 else x - 12)
        self.df['Value'] = pd.to_numeric(self.df['Value'], errors='coerce').fillna(0)

    def define_constants(self):
        self.CORPORATE_ACCOUNTS = [
            'G&A Admin SG&A',
            'Operations SG&A',
            'Shared Service Allocations',
            'G&A Admin Headcount',
            'Depreciation'
        ]

        self.COGS_ACCOUNTS = [
            'Labor COGs',
            'Material COGs',
            'Other COGs'
        ]

        self.SGA_ACCOUNTS = [
            'Marketing SG&A',
            'Sales SG&A',
            'Operations SG&A',
            'Implementation SG&A',
            'G&A Admin SG&A'
        ]

        self.HC_ACCOUNTS = [
            'Sales Headcount',
            'Implementation Headcount',
            'Operations Headcount',
            'Marketing Headcount',
            'G&A Admin Headcount'
        ]

        self.ALL_EXPENSES = (
            self.COGS_ACCOUNTS +
            self.SGA_ACCOUNTS +
            ['Shared Service Allocations', 'Depreciation']
        )

        self.BRANCHES = [f'Branch {i}' for i in range(1, 11)]
        self.ALL_ENTITIES = ['All Branches'] + self.BRANCHES + ['Corporate']

        self.REGION_MAP = {
            'Branch 1': 'G',
            'Branch 2': 'G',
            'Branch 3': 'C',
            'Branch 4': 'B',
            'Branch 5': 'F',
            'Branch 6': 'A',
            'Branch 7': 'A',
            'Branch 8': 'D',
            'Branch 9': 'D',
            'Branch 10': 'E'
        }

        self.EXPENSE_LINES = [
            'Labor COGs',
            'Material COGs',
            'Other COGs',
            'Total COGS',
            'Marketing SG&A',
            'Sales SG&A',
            'Operations SG&A',
            'Implementation SG&A',
            'G&A Admin SG&A',
            'Total SG&A',
            'Allocations',
            'Depreciation'
        ]

        self.IS_LINES = [
            ('Revenue', 'Revenue'),
            ('  Labor COGs', 'Labor COGs'),
            ('  Material COGs', 'Material COGs'),
            ('  Other COGs', 'Other COGs'),
            ('Total COGS', self.COGS_ACCOUNTS),
            ('Gross Profit', 'GROSS_PROFIT'),
            ('  Marketing SG&A', 'Marketing SG&A'),
            ('  Sales SG&A', 'Sales SG&A'),
            ('  Operations SG&A', 'Operations SG&A'),
            ('  Implementation SG&A', 'Implementation SG&A'),
            ('  G&A Admin SG&A', 'G&A Admin SG&A'),
            ('Total SG&A', self.SGA_ACCOUNTS),
            ('Allocations', 'Shared Service Allocations'),
            ('EBITDA', 'EBITDA'),
            ('Depreciation', 'Depreciation'),
            ('EBIT', 'EBIT'),
        ]
    

    def get_data(self, entity, acct, yr, mo=None):
        """Entity kurallarına göre veriyi süzer."""
        acts = acct if isinstance(acct, list) else [acct]

        if entity == 'All Branches':
            # Corporate hesapları sadece corporate hesap listesinde varsa dahil et
            use_branches = (
                self.BRANCHES + ['Corporate']
                if any(a in self.CORPORATE_ACCOUNTS for a in acts)
                else self.BRANCHES
            )
        else:
            use_branches = [entity]

        mask = (
            (self.df['Year'] == yr) &
            (self.df['Account'].isin(acts)) &
            (self.df['Branch'].isin(use_branches))
        )

        if mo:
            mask &= (self.df['Month'] == mo)

        return self.df.loc[mask, 'Value'].sum()

    def calc_ebitda(self, entity, yr, mo=None):
        rev = self.get_data(entity, 'Revenue', yr, mo)
        cogs = self.get_data(entity, self.COGS_ACCOUNTS, yr, mo)
        sga = self.get_data(entity, self.SGA_ACCOUNTS, yr, mo)
        alloc = self.get_data(entity, 'Shared Service Allocations', yr, mo)
        return rev - cogs - sga - alloc

    def calc_ebit(self, entity, yr, mo=None):
        return self.calc_ebitda(entity, yr, mo) - self.get_data(entity, 'Depreciation', yr, mo)

    # -----------------------------
    # RAPOR 1 — IS Comparison
    # -----------------------------
    # -----------------------------
    # RAPOR 1 — IS Comparison (Backend Only)
    # -----------------------------
    def build_report1_is_comparison(self, entity):
        display_lines = [
            ('Revenue',               'Revenue'),
            ('── COGS ──',            None),
            ('  Labor COGs',          'Labor COGs'),
            ('  Material COGs',       'Material COGs'),
            ('  Other COGs',          'Other COGs'),
            ('Total COGS',            self.COGS_ACCOUNTS),
            ('Gross Profit',          'GROSS_PROFIT'),
            ('── SG&A ──',            None),
            ('  Marketing SG&A',      'Marketing SG&A'),
            ('  Sales SG&A',          'Sales SG&A'),
            ('  Operations SG&A',     'Operations SG&A'),
            ('  Implementation SG&A', 'Implementation SG&A'),
            ('  G&A Admin SG&A',      'G&A Admin SG&A'),
            ('Total SG&A',            self.SGA_ACCOUNTS),
            ('Allocations',           'Shared Service Allocations'),
            ('EBITDA',                'EBITDA'),
            ('Depreciation',          'Depreciation'),
            ('EBIT',                  'EBIT'),
        ]

        results = []

        for label, acct in display_lines:
            # Sadece başlık olan satırlar (Frontend'de stil vermek için)
            if acct is None:
                results.append({
                    'Label': label,
                    'Nov_2024': None,
                    'Dec_2024': None,
                    'MoM_$': None,
                    'MoM_%': None,
                    'FY_2023': None,
                    'FY_2024': None,
                    'YoY_$': None,
                    'YoY_%': None,
                    'Is_Header': True
                })
                continue

            # Veri Çekme & Özel Hesaplamalar
            if acct == 'GROSS_PROFIT':
                nov  = self.get_data(entity, 'Revenue', 2024, 11) - self.get_data(entity, self.COGS_ACCOUNTS, 2024, 11)
                dec  = self.get_data(entity, 'Revenue', 2024, 12) - self.get_data(entity, self.COGS_ACCOUNTS, 2024, 12)
                fy23 = self.get_data(entity, 'Revenue', 2023)     - self.get_data(entity, self.COGS_ACCOUNTS, 2023)
                fy24 = self.get_data(entity, 'Revenue', 2024)     - self.get_data(entity, self.COGS_ACCOUNTS, 2024)
            elif acct == 'EBITDA':
                nov = self.calc_ebitda(entity, 2024, 11)
                dec = self.calc_ebitda(entity, 2024, 12)
                fy23 = self.calc_ebitda(entity, 2023)
                fy24 = self.calc_ebitda(entity, 2024)
            elif acct == 'EBIT':
                nov = self.calc_ebit(entity, 2024, 11)
                dec = self.calc_ebit(entity, 2024, 12)
                fy23 = self.calc_ebit(entity, 2023)
                fy24 = self.calc_ebit(entity, 2024)
            else:
                nov = self.get_data(entity, acct, 2024, 11)
                dec = self.get_data(entity, acct, 2024, 12)
                fy23 = self.get_data(entity, acct, 2023)
                fy24 = self.get_data(entity, acct, 2024)

            # Varyans (Değişim) Hesaplamaları
            # Gider kalemlerindeki artışlar negatif (kötü) olarak işaretlenir
            sign = -1 if label.strip() in self.EXPENSE_LINES else 1
            
            mom_d = sign * (dec - nov)
            mom_p = sign * (dec - nov) / abs(nov) if nov != 0 else None
            
            yoy_d = sign * (fy24 - fy23)
            yoy_p = sign * (fy24 - fy23) / abs(fy23) if fy23 != 0 else None

            # Sonuçları ekle
            results.append({
                'Label': label,
                'Nov_2024': nov,
                'Dec_2024': dec,
                'MoM_$': mom_d,
                'MoM_%': mom_p,
                'FY_2023': fy23,
                'FY_2024': fy24,
                'YoY_$': yoy_d,
                'YoY_%': yoy_p,
                'Is_Header': False
            })

        return pd.DataFrame(results)

    # -----------------------------
    # RAPOR 2 — Monthly IS
    # -----------------------------
    # -----------------------------
    # RAPOR 2 — Monthly Trended IS (Backend Only)
    # -----------------------------
    def build_report2_trended_is(self, entity, period='Full Year Monthly'):
        display_lines = [
            ('Revenue',               'Revenue'),
            ('── COGS ──',            None),
            ('  Labor COGs',          'Labor COGs'),
            ('  Material COGs',       'Material COGs'),
            ('  Other COGs',          'Other COGs'),
            ('Total COGS',            self.COGS_ACCOUNTS),
            ('Gross Profit',          'GROSS_PROFIT'),
            ('── SG&A ──',            None),
            ('  Marketing SG&A',      'Marketing SG&A'),
            ('  Sales SG&A',          'Sales SG&A'),
            ('  Operations SG&A',     'Operations SG&A'),
            ('  Implementation SG&A', 'Implementation SG&A'),
            ('  G&A Admin SG&A',      'G&A Admin SG&A'),
            ('Total SG&A',            self.SGA_ACCOUNTS),
            ('Allocations',           'Shared Service Allocations'),
            ('EBITDA',                'EBITDA'),
            ('Depreciation',          'Depreciation'),
            ('EBIT',                  'EBIT'),
        ]

        # Seçilen periyoda göre ayları belirle
        if period == 'Q1 (Jan-Mar)': months = [1, 2, 3]
        elif period == 'Q2 (Apr-Jun)': months = [4, 5, 6]
        elif period == 'Q3 (Jul-Sep)': months = [7, 8, 9]
        elif period == 'Q4 (Oct-Dec)': months = [10, 11, 12]
        elif period == 'Annual Only': months = []
        else: months = list(range(1, 13)) # Full Year Monthly

        results = []

        for label, acct in display_lines:
            row = {'Label': label}
            
            # Sadece başlık olan satırlar (Frontend'de CSS için)
            if acct is None:
                row['Is_Header'] = True
                for yr in [2023, 2024]:
                    for m in months:
                        row[f'{yr}_M{m}'] = None
                    row[f'FY_{yr}'] = None
                results.append(row)
                continue

            row['Is_Header'] = False

            # Veri çekme ve hesaplama için yardımcı fonksiyon
            def get_val(yr, mo=None):
                if acct == 'GROSS_PROFIT':
                    return self.get_data(entity, 'Revenue', yr, mo) - self.get_data(entity, self.COGS_ACCOUNTS, yr, mo)
                elif acct == 'EBITDA':
                    return self.calc_ebitda(entity, yr, mo)
                elif acct == 'EBIT':
                    return self.calc_ebit(entity, yr, mo)
                else:
                    return self.get_data(entity, acct, yr, mo)

            # 2023 ve 2024 verilerini döngüyle çekip satıra ekle
            for yr in [2023, 2024]:
                for m in months:
                    row[f'{yr}_M{m}'] = get_val(yr, m)
                row[f'FY_{yr}'] = get_val(yr)

            results.append(row)

        # DataFrame'i oluştur ve kolon sırasını frontend'in beklediği mantığa göre diz
        cols = ['Label']
        for m in months: cols.append(f'2023_M{m}')
        cols.append('FY_2023')
        for m in months: cols.append(f'2024_M{m}')
        cols.append('FY_2024')
        cols.append('Is_Header')

        return pd.DataFrame(results)[cols]
    # -----------------------------
    # RAPOR 3 — Rankings
    # -----------------------------
    # -----------------------------
    # RAPOR 3 — Rankings (Backend Only)
    # -----------------------------
    def build_report3_rankings(self):
        results = []

        for b in self.BRANCHES:
            # 1. YoY Sales Growth
            rev23 = self.get_data(b, 'Revenue', 2023)
            rev24 = self.get_data(b, 'Revenue', 2024)
            yoy_sales = (rev24 - rev23) / rev23 if rev23 != 0 else None

            # 2. 2024 ARPU
            ru24 = self.get_data(b, 'Revenue (Units)', 2024)
            arpu = rev24 / ru24 if ru24 != 0 else None

            # 3. Sales $ per Headcount
            sales_hc_avg = self.get_data(b, 'Sales Headcount', 2024) / 12
            sales_hc = rev24 / sales_hc_avg if sales_hc_avg != 0 else None

            # 4. YoY Gross Margin Growth
            cogs23 = self.get_data(b, self.COGS_ACCOUNTS, 2023)
            cogs24 = self.get_data(b, self.COGS_ACCOUNTS, 2024)
            gm23 = rev23 - cogs23
            gm24 = rev24 - cogs24
            yoy_gm = (gm24 - gm23) / abs(gm23) if gm23 != 0 else None

            # 5. YoY Order Growth
            ord23 = self.get_data(b, 'Orders (Units)', 2023)
            ord24 = self.get_data(b, 'Orders (Units)', 2024)
            yoy_ord = (ord24 - ord23) / ord23 if ord23 != 0 else None

            results.append({
                'Branch': b,
                'Region': self.REGION_MAP.get(b, 'Unknown'),
                'YoY_Sales_Growth': yoy_sales,
                'ARPU_2024': arpu,
                'Sales_per_HC': sales_hc,
                'YoY_GM_Growth': yoy_gm,
                'YoY_Order_Growth': yoy_ord
            })

        rdf = pd.DataFrame(results)

        # ── Sıralama (Ranking) Hesaplamaları ──
        metrics = [
            'YoY_Sales_Growth',
            'ARPU_2024',
            'Sales_per_HC',
            'YoY_GM_Growth',
            'YoY_Order_Growth'
        ]

        # Her bir metrik için şubeleri sırala (En yüksek değer = 1. sıra)
        for m in metrics:
            rdf[f'Rank_{m}'] = rdf[m].rank(ascending=False, na_option='bottom').astype(int)

        # Ortalama sıralama (Avg Rank)
        rank_cols = [f'Rank_{m}' for m in metrics]
        rdf['Avg_Rank'] = rdf[rank_cols].mean(axis=1)

        # En düşük ortalamaya sahip şube genel 1. olacak şekilde tabloyu sırala
        rdf = rdf.sort_values('Avg_Rank').reset_index(drop=True)
        
        # Overall Rank kolonunu tablonun en başına ekle (1, 2, 3...)
        rdf.insert(0, 'Overall_Rank', rdf.index + 1)

        return rdf
    # -----------------------------
    # RAPOR 4 — Regional Comparison
    # -----------------------------
    # -----------------------------
    # RAPOR 4 — Regional Comparison (Backend Only)
    # -----------------------------
    def build_report4_regional(self):
        # Bölgeleri ve içindeki şubeleri grupla
        regions = sorted(list(set(self.REGION_MAP.values())))
        res = []
        
        # Şirket geneli toplamları tutmak için
        grand_total = {
            2023: {'rev': 0, 'sga': 0, 'alloc': 0, 'ebd': 0},
            2024: {'rev': 0, 'sga': 0, 'alloc': 0, 'ebd': 0}
        }

        def get_yoy(new, old):
            return (new - old) / abs(old) if old != 0 else None

        for r in regions:
            branches_in_region = [b for b, reg in self.REGION_MAP.items() if reg == r]
            
            # Bölge toplamlarını hesapla
            r_rev23 = sum(self.get_data(b, 'Revenue', 2023) for b in branches_in_region)
            r_rev24 = sum(self.get_data(b, 'Revenue', 2024) for b in branches_in_region)
            r_sga23 = sum(self.get_data(b, self.SGA_ACCOUNTS, 2023) for b in branches_in_region)
            r_sga24 = sum(self.get_data(b, self.SGA_ACCOUNTS, 2024) for b in branches_in_region)
            r_alc23 = sum(self.get_data(b, 'Shared Service Allocations', 2023) for b in branches_in_region)
            r_alc24 = sum(self.get_data(b, 'Shared Service Allocations', 2024) for b in branches_in_region)
            r_ebd23 = sum(self.calc_ebitda(b, 2023) for b in branches_in_region)
            r_ebd24 = sum(self.calc_ebitda(b, 2024) for b in branches_in_region)

            # Bölge satırını ekle
            res.append({
                'Level': 'Region',
                'Name': f'Region {r}',
                'Rev_23': r_rev23, 'Rev_24': r_rev24, 'Rev_YoY': get_yoy(r_rev24, r_rev23),
                'SGA_23': r_sga23, 'SGA_24': r_sga24, 'SGA_YoY': get_yoy(r_sga24, r_sga23),
                'Alc_23': r_alc23, 'Alc_24': r_alc24, 'Alc_YoY': get_yoy(r_alc24, r_alc23),
                'EBD_23': r_ebd23, 'EBD_24': r_ebd24, 'EBD_YoY': get_yoy(r_ebd24, r_ebd23)
            })

            # Şube detay satırlarını ekle
            for b in branches_in_region:
                b_rev23, b_rev24 = self.get_data(b, 'Revenue', 2023), self.get_data(b, 'Revenue', 2024)
                b_sga23, b_sga24 = self.get_data(b, self.SGA_ACCOUNTS, 2023), self.get_data(b, self.SGA_ACCOUNTS, 2024)
                b_alc23, b_alc24 = self.get_data(b, 'Shared Service Allocations', 2023), self.get_data(b, 'Shared Service Allocations', 2024)
                b_ebd23, b_ebd24 = self.calc_ebitda(b, 2023), self.calc_ebitda(b, 2024)

                res.append({
                    'Level': 'Branch',
                    'Name': b,
                    'Rev_23': b_rev23, 'Rev_24': b_rev24, 'Rev_YoY': get_yoy(b_rev24, b_rev23),
                    'SGA_23': b_sga23, 'SGA_24': b_sga24, 'SGA_YoY': get_yoy(b_sga24, b_sga23),
                    'Alc_23': b_alc23, 'Alc_24': b_alc24, 'Alc_YoY': get_yoy(b_alc24, b_alc23),
                    'EBD_23': b_ebd23, 'EBD_24': b_ebd24, 'EBD_YoY': get_yoy(b_ebd24, b_ebd23)
                })

            # Grand total için topla
            for yr, rev, sga, alc, ebd in [(2023, r_rev23, r_sga23, r_alc23, r_ebd23), (2024, r_rev24, r_sga24, r_alc24, r_ebd24)]:
                grand_total[yr]['rev'] += rev
                grand_total[yr]['sga'] += sga
                grand_total[yr]['alloc'] += alc
                grand_total[yr]['ebd'] += ebd

        # Corporate satırı
        c_sga23, c_sga24 = self.get_data('Corporate', self.SGA_ACCOUNTS, 2023), self.get_data('Corporate', self.SGA_ACCOUNTS, 2024)
        c_alc23, c_alc24 = self.get_data('Corporate', 'Shared Service Allocations', 2023), self.get_data('Corporate', 'Shared Service Allocations', 2024)
        c_ebd23, c_ebd24 = self.calc_ebitda('Corporate', 2023), self.calc_ebitda('Corporate', 2024)

        res.append({
            'Level': 'Corporate',
            'Name': 'Corporate',
            'Rev_23': 0, 'Rev_24': 0, 'Rev_YoY': 0,
            'SGA_23': c_sga23, 'SGA_24': c_sga24, 'SGA_YoY': get_yoy(c_sga24, c_sga23),
            'Alc_23': c_alc23, 'Alc_24': c_alc24, 'Alc_YoY': get_yoy(c_alc24, c_alc23),
            'EBD_23': c_ebd23, 'EBD_24': c_ebd24, 'EBD_YoY': get_yoy(c_ebd24, c_ebd23)
        })

        # Total Company satırı (Tahsisatlar düşülmüş net EBITDA)
        t_ebd23 = grand_total[2023]['ebd'] + c_ebd23 - c_alc23
        t_ebd24 = grand_total[2024]['ebd'] + c_ebd24 - c_alc24

        res.append({
            'Level': 'Total',
            'Name': '🏢 Total Company',
            'Rev_23': grand_total[2023]['rev'], 'Rev_24': grand_total[2024]['rev'], 'Rev_YoY': get_yoy(grand_total[2024]['rev'], grand_total[2023]['rev']),
            'SGA_23': grand_total[2023]['sga'] + c_sga23, 'SGA_24': grand_total[2024]['sga'] + c_sga24, 'SGA_YoY': get_yoy(grand_total[2024]['sga'] + c_sga24, grand_total[2023]['sga'] + c_sga23),
            'Alc_23': grand_total[2023]['alloc'] + c_alc23, 'Alc_24': grand_total[2024]['alloc'] + c_alc24, 'Alc_YoY': get_yoy(grand_total[2024]['alloc'] + c_alc24, grand_total[2023]['alloc'] + c_alc23),
            'EBD_23': t_ebd23, 'EBD_24': t_ebd24, 'EBD_YoY': get_yoy(t_ebd24, t_ebd23)
        })

        return pd.DataFrame(res)

    # -----------------------------
    # RAPOR 5 — Metrics
    # -----------------------------
    # -----------------------------
    # RAPOR 5 — Metrics Dashboard (Backend Only)
    # -----------------------------
    def build_report5_metrics(self, entity):
        import calendar
        
        # 2024 yılı ay bazlı gün sayıları (Artık yıl kontrolü: 366 gün)
        days_in_months = [calendar.monthrange(2024, m)[1] for m in range(1, 13)]
        
        # Entity kuralları
        if entity == 'All Branches':
            branches = self.BRANCHES + ['Corporate']
            branches_no_corp = self.BRANCHES
        elif entity == 'Corporate':
            branches = ['Corporate']
            branches_no_corp = ['Corporate']
        else:
            branches = [entity]
            branches_no_corp = [entity]

        # Yardımcı Veri Çekme Fonksiyonu
        def get_m(acct, m):
            acts = acct if isinstance(acct, list) else [acct]
            # Kurumsal hesaplar kontrolü
            use = branches if any(a in self.CORPORATE_ACCOUNTS for a in acts) else branches_no_corp
            return sum(self.get_data(b, a, 2024, m) for b in use for a in acts)

        def get_a(acct):
            acts = acct if isinstance(acct, list) else [acct]
            use = branches if any(a in self.CORPORATE_ACCOUNTS for a in acts) else branches_no_corp
            return sum(self.get_data(b, a, 2024) for b in use for a in acts)

        # --- AYLIK HESAPLAMALAR ---
        res_rows = []
        
        # Tanımlı Metrik Grupları
        metrics_def = [
            ('Efficiency', [
                ('Impl. Hours per Impl. HC', lambda m: get_m('Implementation Hours', m) / get_m('Implementation Headcount', m) if get_m('Implementation Headcount', m) else None),
                ('Revenue per Direct Labor HC', lambda m: get_m('Revenue', m) / get_m('Implementation Headcount', m) if get_m('Implementation Headcount', m) else None),
                ('Revenue per Sales HC', lambda m: get_m('Revenue', m) / get_m('Sales Headcount', m) if get_m('Sales Headcount', m) else None),
                ('Expenses per Total HC', lambda m: sum(get_m(a, m) for a in self.ALL_EXPENSES) / sum(get_m(a, m) for a in self.HC_ACCOUNTS) if sum(get_m(a, m) for a in self.HC_ACCOUNTS) else None),
            ]),
            ('Volume', [
                ('Backlog Turn Rate %', lambda m: get_m('Revenue (Units)', m) / get_m('Project Backlog (Units)', m) if get_m('Project Backlog (Units)', m) else None),
                ('Backlog Days', lambda m: get_m('Project Backlog (Units)', m) / (get_m('Revenue (Units)', m) / days_in_months[m-1]) if get_m('Revenue (Units)', m) else None),
            ]),
            ('Profitability', [
                ('ARPU', lambda m: get_m('Revenue', m) / get_m('Revenue (Units)', m) if get_m('Revenue (Units)', m) else None),
                ('COGS per Rev Unit', lambda m: sum(get_m(a, m) for a in self.COGS_ACCOUNTS) / get_m('Revenue (Units)', m) if get_m('Revenue (Units)', m) else None),
                ('EBITDA % of Revenue', lambda m: sum(self.calc_ebitda(b, 2024, m) for b in branches_no_corp) / get_m('Revenue', m) if get_m('Revenue', m) else None),
            ])
        ]

        for section, metrics in metrics_def:
            # Sektör Başlığı
            res_rows.append({'Metric': f'── {section.upper()} ──', 'Is_Section': True})
            
            for label, func in metrics:
                row = {'Metric': label, 'Is_Section': False}
                # 12 Ay
                for m in range(1, 13):
                    row[f'M{m}'] = func(m)
                
                # --- YILLIK HESAPLAMA (Annual) ---
                if label == 'Impl. Hours per Impl. HC':
                    row['Annual'] = get_a('Implementation Hours') / (get_a('Implementation Headcount') / 12) if get_a('Implementation Headcount') else None
                elif label == 'Revenue per Direct Labor HC':
                    row['Annual'] = get_a('Revenue') / (get_a('Implementation Headcount') / 12) if get_a('Implementation Headcount') else None
                elif label == 'Revenue per Sales HC':
                    row['Annual'] = get_a('Revenue') / (get_a('Sales Headcount') / 12) if get_a('Sales Headcount') else None
                elif label == 'Expenses per Total HC':
                    total_exp = sum(get_a(a) for a in self.ALL_EXPENSES)
                    avg_hc = sum(get_a(a) for a in self.HC_ACCOUNTS) / 12
                    row['Annual'] = total_exp / avg_hc if avg_hc else None
                elif label == 'Backlog Turn Rate %':
                    row['Annual'] = get_a('Revenue (Units)') / (get_a('Project Backlog (Units)') / 12) if get_a('Project Backlog (Units)') else None
                elif label == 'Backlog Days':
                    avg_backlog = get_a('Project Backlog (Units)') / 12
                    row['Annual'] = avg_backlog / (get_a('Revenue (Units)') / 366) if get_a('Revenue (Units)') else None
                elif label == 'ARPU':
                    row['Annual'] = get_a('Revenue') / get_a('Revenue (Units)') if get_a('Revenue (Units)') else None
                elif label == 'COGS per Rev Unit':
                    row['Annual'] = sum(get_a(a) for a in self.COGS_ACCOUNTS) / get_a('Revenue (Units)') if get_a('Revenue (Units)') else None
                elif label == 'EBITDA % of Revenue':
                    total_ebd = sum(self.calc_ebitda(b, 2024) for b in branches_no_corp)
                    row['Annual'] = total_ebd / get_a('Revenue') if get_a('Revenue') else None
                
                res_rows.append(row)

        return pd.DataFrame(res_rows)
# 1. Excel dosyanızın adını buraya yazın (class'ın dışında, en sonda)
# backend.py dosyasının en sonu böyle olmalı:
if __name__ == "__main__":
    file_path = r'C:\Users\tasan\OneDrive\Masaüstü\GEN_AI_PRJCT\Raw Data for Group.xlsx'
    engine = FinancialDataEngine(file_path)
    print("Veri başarıyla yüklendi.")