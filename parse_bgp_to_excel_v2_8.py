# parse_bgp_to_excel.py (v2.8, kompatibel Py3.8+)
# Fokus perbaikan v2.8:
# - PERBAIKAN PARSER LOG GABUNGAN: header "HOST:" kini fleksibel (boleh ada timestamp "| ...")
#   dan alamat IP opsional. Ini mencegah kasus Excel kosong karena section tidak terdeteksi.
# - FALLBACK TAMBAHAN: jika header HOST tidak ada, tapi ada marker ">>> show bgp neighbor",
#   parser akan memproses seluruh log sebagai satu host "LOG" (agar tetap ada data).
# - Semua peningkatan styling v2.7 tetap dipertahankan (header berwarna, freeze, autofilter,
#   auto-fit, angka ribuan, Delta + color scale & icon set, data bars, table style, Summary & Dashboard).
"""
Parser Junos "show bgp neighbor" → Excel dengan kolom:
 Peer IP, Deskripsi, IPv4 Received, IPv4 Advertised, IPv6 Received,
 IPv6 Advertised, Total Received, Total Advertised, State, Sumber
Fitur:
- Baca log gabungan (--log). Kalau header 'HOST:' tak ada, fallback ke folder logs\\\\YYYYMMDD (file per-host).
- Robust ke variasi baris 'Peer:':
 "Peer: <peer>[+port] AS <asn> Local: <local>[+port] [AS <asn_local>]"
- Marker '>>>' opsional; bila tak ada, parser gunakan seluruh blok & cari 'Peer:' langsung.
Exit codes:
 0 = sukses
 2 = file log tidak ditemukan
 3 = tidak ada data BGP yang terdeteksi dari log
 4 = dependency (pandas/openpyxl/xlsxwriter) tidak tersedia
 5 = gagal menulis Excel
"""
import argparse
import os
import re
import sys
from datetime import datetime

try:
    import pandas as pd
except Exception:
    sys.stderr.write("[ERROR] Paket 'pandas' belum terpasang. Instal: py -3 -m pip install --user pandas\n")
    sys.exit(4)

# --- Regex yang lebih toleran terhadap variasi log gabungan ---
# Contoh baris dari SecureCRT:
#   -------------------------------
#   HOST: R1.BDS.RR-INET.1  (61.5.13.81)  |  1/30/2026 09:12:33 AM
#   -------------------------------
# atau tanpa IP: "HOST: R1.BDS.RR-INET.1 | 1/30/2026 ..."
HOST_HEADER_RE = re.compile(r'^HOST:\s*(?P<name>.+?)(?:\s*\((?P<addr>.+?)\))?(?:\s*\|\s*.*)?$')
CMD_MARK_RE    = re.compile(r'^>>>\s*(?P<cmd>.+?)\s*$')
PEER_LINE_RE   = re.compile(
    r'^\s*Peer:\s*(?P<peer>\S+?)' \
    r'(?:\+(?P<peer_port>\d+))?\s+AS\s+(?P<asn>\d+)' \
    r'(?:\s+Local:\s*(?P<local>\S+?)(?:\+(?P<lport>\d+))?)?' \
    r'(?:\s+AS\s+(?P<asn_local>\d+))?\s*$',
    re.IGNORECASE,
)
TYPE_STATE_RE  = re.compile(
    r'^\s*Type:\s*(?P<type>\S+)\s+State:\s*(?P<state>\S+)(?:\s+(?P<uptime>\d+)\s+seconds)?',
    re.IGNORECASE,
)
DESC_RE        = re.compile(r'^\s*Description:\s*(?P<desc>.*)$', re.IGNORECASE)
V4_TABLE_RE    = re.compile(r'^\s*Table\s+bgp\.l3vpn\.0\b', re.IGNORECASE)
V6_TABLE_RE    = re.compile(r'^\s*Table\s+bgp\.l3vpn-?inet6\.0\b', re.IGNORECASE)
RECEIVED_RE    = re.compile(r'Received\s+prefixes:\s+(?P<num>\d+)', re.IGNORECASE)
ADVERTISED_RE  = re.compile(r'Advertised\s+prefixes:\s+(?P<num>\d+)', re.IGNORECASE)


def split_by_hosts_via_host_header(log_text: str):
    """Pisahkan log gabungan menjadi {host_name: teks_host}.
    Pencarian host berbasis baris yang diawali 'HOST:' dan toleran terhadap tambahan timestamp.
    """
    sections = {}
    current_host = None
    capture = []
    for line in log_text.splitlines():
        m = HOST_HEADER_RE.match(line or "")
        if m:
            # simpan section sebelumnya
            if current_host is not None and capture:
                sections[current_host] = "\n".join(capture).strip()
                capture = []
            # set host aktif
            current_host = (m.group("name") or '').strip()
        else:
            if current_host is not None:
                capture.append(line)
    if current_host is not None and capture:
        sections[current_host] = "\n".join(capture).strip()
    return sections


def extract_cmd_block(text: str, cmd_prefix: str):
    """Ambil blok output yang mengikuti marker ">>> {cmd_prefix}..." hingga marker berikutnya/EOF."""
    lines = text.splitlines()
    out_lines = []
    capturing = False
    for line in lines:
        m = CMD_MARK_RE.match(line)
        if m:
            cmd = m.group("cmd").strip().lower()
            if cmd.startswith(cmd_prefix.lower()):
                capturing = True
                continue
            else:
                if capturing:
                    break
                else:
                    continue
        else:
            if capturing:
                out_lines.append(line)
    return "\n".join(out_lines).strip()


def parse_junos_show_bgp_neighbor(text: str, host_display: str):
    results = []
    blocks = []
    cur = []
    for raw in text.splitlines():
        if raw.strip().lower().startswith("peer:"):
            if cur:
                blocks.append(cur)
            cur = [raw]
        else:
            if cur:
                cur.append(raw)
    if cur:
        blocks.append(cur)

    for blk in blocks:
        entry = {
            'Peer IP': '',
            'Deskripsi': '',
            'IPv4 Received': 0,
            'IPv4 Advertised': 0,
            'IPv6 Received': 0,
            'IPv6 Advertised': 0,
            'Total Received': 0,
            'Total Advertised': 0,
            'State': '',
            'Sumber': host_display,
        }
        in_v4 = in_v6 = False
        v4_recv_found = v4_adv_found = v6_recv_found = v6_adv_found = False
        for line in blk:
            m1 = PEER_LINE_RE.match(line)
            if m1:
                entry['Peer IP'] = (m1.group('peer') or '').strip()
                continue
            m2 = TYPE_STATE_RE.match(line)
            if m2:
                entry['State'] = (m2.group('state') or '').title()
                continue
            m3 = DESC_RE.match(line)
            if m3:
                entry['Deskripsi'] = (m3.group('desc') or '').strip()
                continue
            if V4_TABLE_RE.match(line):
                in_v4, in_v6 = True, False
                continue
            if V6_TABLE_RE.match(line):
                in_v4, in_v6 = False, True
                continue
            if in_v4:
                mr = RECEIVED_RE.search(line)
                if mr and not v4_recv_found:
                    try:
                        entry['IPv4 Received'] = int(mr.group('num'))
                        v4_recv_found = True
                    except Exception:
                        pass
                ma = ADVERTISED_RE.search(line)
                if ma and not v4_adv_found:
                    try:
                        entry['IPv4 Advertised'] = int(ma.group('num'))
                        v4_adv_found = True
                    except Exception:
                        pass
            if in_v6:
                mr6 = RECEIVED_RE.search(line)
                if mr6 and not v6_recv_found:
                    try:
                        entry['IPv6 Received'] = int(mr6.group('num'))
                        v6_recv_found = True
                    except Exception:
                        pass
                ma6 = ADVERTISED_RE.search(line)
                if ma6 and not v6_adv_found:
                    try:
                        entry['IPv6 Advertised'] = int(ma6.group('num'))
                        v6_adv_found = True
                    except Exception:
                        pass
        entry['Total Received']    = (entry['IPv4 Received'] or 0) + (entry['IPv6 Received'] or 0)
        entry['Total Advertised']  = (entry['IPv4 Advertised'] or 0) + (entry['IPv6 Advertised'] or 0)
        if entry['Peer IP']:
            results.append(entry)
    return results

# ================== Styling (xlsxwriter) ==================

def _auto_fit_columns(df, worksheet, workbook, start_col=0):
    max_widths = []
    for col in df.columns:
        header_len = len(str(col))
        data_len = df[col].astype(str).map(len).max() if not df.empty else 0
        max_widths.append(min(max(header_len, data_len) + 2, 60))
    for idx, width in enumerate(max_widths):
        worksheet.set_column(start_col + idx, start_col + idx, width)


def _add_summary_sheet(data_per_host, writer, engine, theme):
    import pandas as pd
    rows = []
    for _, items in data_per_host.items():
        rows.extend(items)
    if not rows:
        return None

    df = pd.DataFrame(rows)
    num_cols = ['IPv4 Received','IPv4 Advertised','IPv6 Received','IPv6 Advertised','Total Received','Total Advertised']
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
    df['Delta (Recv-Adv)'] = df['Total Received'] - df['Total Advertised']

    summary = (df.groupby('Sumber', as_index=False)[num_cols + ['Delta (Recv-Adv)']].sum()
              ).sort_values('Total Received', ascending=False)

    sheet = 'Summary'
    summary.to_excel(writer, index=False, sheet_name=sheet)

    if engine == 'xlsxwriter':
        wb = writer.book
        ws = writer.sheets[sheet]
        header_fmt = wb.add_format({'bold': True, 'bg_color': theme['header_bg'], 'font_color': 'white', 'border': 0})
        num_fmt    = wb.add_format({'num_format': '#,##0', 'align': 'right'})
        left_fmt   = wb.add_format({'align': 'left'})
        ws.set_row(0, None, header_fmt)
        ws.autofilter(0, 0, len(summary), len(summary.columns)-1)
        ws.freeze_panes(1, 1)
        _auto_fit_columns(summary, ws, wb)
        for j, col in enumerate(summary.columns):
            ws.set_column(j, j, None, num_fmt if col != 'Sumber' else left_fmt)
        # Grand Total
        last_row = len(summary) + 1
        ws.write(last_row, 0, 'Grand Total', wb.add_format({'bold': True}))
        for j in range(1, len(summary.columns)):
            col_letter = chr(ord('A') + j)
            ws.write_formula(last_row, j, f'=SUBTOTAL(9,{col_letter}2:{col_letter}{len(summary)+1})', num_fmt)
        ws.set_tab_color(theme['tab'])
    return summary


def _add_dashboard(writer, engine, summary_df, state_counts, theme):
    if engine != 'xlsxwriter' or summary_df is None:
        return
    wb = writer.book
    sheet = 'Dashboard'
    ws = wb.add_worksheet(sheet)
    title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': theme['title']})
    ws.write('A1', 'BGP Neighbors Dashboard', title_fmt)
    ws.set_tab_color(theme['tab'])

    # Top-10 by Total Received
    top = summary_df[['Sumber','Total Received']].head(10).copy()
    start_row = 3
    ws.write(start_row, 0, 'Top-10 Sumber by Total Received', wb.add_format({'bold': True}))
    ws.write(start_row+1, 0, 'Sumber', wb.add_format({'bold': True, 'bg_color': theme['header_bg'], 'font_color': 'white'}))
    ws.write(start_row+1, 1, 'Total Received', wb.add_format({'bold': True, 'bg_color': theme['header_bg'], 'font_color': 'white'}))
    for i, (_, row) in enumerate(top.iterrows(), start=start_row+2):
        ws.write(i, 0, row['Sumber'])
        ws.write_number(i, 1, int(row['Total Received']))
    chart1 = wb.add_chart({'type': 'column'})
    chart1.add_series({
        'name':       'Total Received',
        'categories': [sheet, start_row+2, 0, start_row+1+len(top), 0],
        'values':     [sheet, start_row+2, 1, start_row+1+len(top), 1],
        'fill':       {'color': theme['bar']},
    })
    chart1.set_title({'name': 'Top-10 Total Received'})
    chart1.set_legend({'position': 'none'})
    chart1.set_y_axis({'major_gridlines': {'visible': False}})
    ws.insert_chart(start_row+1, 3, chart1, {'x_scale': 1.3, 'y_scale': 1.1})

    # Donut: State distribution
    kpi_row = start_row + 2 + len(top) + 3
    ws.write(kpi_row, 0, 'Overall State Distribution', wb.add_format({'bold': True}))
    ws.write_row(kpi_row+1, 0, ['State', 'Count'], wb.add_format({'bold': True, 'bg_color': theme['header_bg'], 'font_color': 'white'}))
    st_items = list(state_counts.items())
    for idx, (st, cnt) in enumerate(st_items, start=kpi_row+2):
        ws.write(idx, 0, st)
        ws.write_number(idx, 1, int(cnt))
    chart2 = wb.add_chart({'type': 'doughnut'})
    chart2.add_series({
        'name': 'States',
        'categories': [sheet, kpi_row+2, 0, kpi_row+1+len(st_items), 0],
        'values':     [sheet, kpi_row+2, 1, kpi_row+1+len(st_items), 1],
        'points': [
            {'fill': {'color': '#70AD47'}},
            {'fill': {'color': '#FFC000'}},
            {'fill': {'color': '#FF0000'}},
            {'fill': {'color': '#A5A5A5'}},
        ],
    })
    chart2.set_title({'name': 'States'})
    chart2.set_hole_size(60)
    ws.insert_chart(kpi_row, 3, chart2, {'x_scale': 1.2, 'y_scale': 1.2})

    # Stacked: Received vs Advertised (Top 10)
    top2 = summary_df[['Sumber','Total Received','Total Advertised']].head(10).copy()
    sec_row = kpi_row + 2 + len(st_items) + 3
    ws.write(sec_row, 0, 'Top-10 Received vs Advertised', wb.add_format({'bold': True}))
    ws.write_row(sec_row+1, 0, ['Sumber','Total Received','Total Advertised'], wb.add_format({'bold': True, 'bg_color': theme['header_bg'], 'font_color': 'white'}))
    for i, (_, row) in enumerate(top2.iterrows(), start=sec_row+2):
        ws.write(i, 0, row['Sumber'])
        ws.write_number(i, 1, int(row['Total Received']))
        ws.write_number(i, 2, int(row['Total Advertised']))
    chart3 = wb.add_chart({'type': 'column', 'subtype': 'stacked'})
    chart3.add_series({
        'name': 'Total Received',
        'categories': [sheet, sec_row+2, 0, sec_row+1+len(top2), 0],
        'values':     [sheet, sec_row+2, 1, sec_row+1+len(top2), 1],
        'fill': {'color': theme['bar']}
    })
    chart3.add_series({
        'name': 'Total Advertised',
        'categories': [sheet, sec_row+2, 0, sec_row+1+len(top2), 0],
        'values':     [sheet, sec_row+2, 2, sec_row+1+len(top2), 2],
        'fill': {'color': '#ED7D31'}
    })
    chart3.set_title({'name': 'Received vs Advertised (Top 10)'})
    chart3.set_y_axis({'major_gridlines': {'visible': False}})
    ws.insert_chart(sec_row+1, 3, chart3, {'x_scale': 1.3, 'y_scale': 1.1})



def to_excel(data_per_host, out_path, mode='per-host', engine_pref='xlsxwriter'):
    cols = ['Peer IP','Deskripsi','IPv4 Received','IPv4 Advertised','IPv6 Received','IPv6 Advertised',
            'Total Received','Total Advertised','State','Sumber']

    theme = {
        'header_bg': '#1F4E78',
        'tab': '#5B9BD5',
        'title': '#1F4E78',
        'bar': '#5B9BD5',
    }

    engine = None
    if engine_pref == 'xlsxwriter':
        try:
            import xlsxwriter  # noqa
            engine = 'xlsxwriter'
        except Exception:
            pass
    if engine is None:
        try:
            import openpyxl  # noqa
            engine = 'openpyxl'
        except Exception:
            pass
    if engine is None:
        sys.stderr.write("[ERROR] Engine Excel tidak tersedia. Instal: xlsxwriter atau openpyxl\n")
        sys.exit(4)

    with pd.ExcelWriter(out_path, engine=engine) as writer:
        all_rows = []
        if mode.lower() == 'single':
            for _, rows_list in data_per_host.items():
                all_rows.extend(rows_list)
            df = pd.DataFrame(all_rows, columns=cols)
            if not df.empty:
                df['Delta (Recv-Adv)'] = (df['Total Received'] - df['Total Advertised']).astype(int)
            else:
                df['Delta (Recv-Adv)'] = pd.Series(dtype=int)
            sheet_name = 'BGP_Neighbors'
            df.to_excel(writer, index=False, sheet_name=sheet_name)

            if engine == 'xlsxwriter':
                wb  = writer.book
                ws  = writer.sheets[sheet_name]
                header_fmt = wb.add_format({'bold': True, 'bg_color': theme['header_bg'], 'font_color': 'white'})
                right_num  = wb.add_format({'num_format': '#,##0', 'align': 'right'})
                left_fmt   = wb.add_format({'align': 'left'})
                ws.set_row(0, None, header_fmt)
                ws.freeze_panes(1, 1)
                _auto_fit_columns(df, ws, wb)
                numeric_cols = ['IPv4 Received','IPv4 Advertised','IPv6 Received','IPv6 Advertised','Total Received','Total Advertised','Delta (Recv-Adv)']
                for j, col in enumerate(df.columns):
                    ws.set_column(j, j, None, right_num if col in numeric_cols else left_fmt)
                if 'State' in df.columns:
                    sidx = df.columns.get_loc('State')
                    ws.conditional_format(1, sidx, len(df), sidx, {
                        'type': 'formula', 'criteria': '=UPPER($%s2)="ESTABLISHED"' % chr(ord('A')+sidx),
                        'format': wb.add_format({'font_color': '#006100', 'bg_color': '#C6EFCE'})
                    })
                    ws.conditional_format(1, sidx, len(df), sidx, {
                        'type': 'formula', 'criteria': '=OR(UPPER($%s2)="IDLE",UPPER($%s2)="ACTIVE",UPPER($%s2)="CONNECT")' % (chr(ord('A')+sidx),chr(ord('A')+sidx),chr(ord('A')+sidx)),
                        'format': wb.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE'})
                    })
                for col_name in ['Total Received','Total Advertised']:
                    if col_name in df.columns:
                        cidx = df.columns.get_loc(col_name)
                        ws.conditional_format(1, cidx, len(df), cidx, {'type': 'data_bar', 'bar_color': theme['bar']})
                if 'Delta (Recv-Adv)' in df.columns:
                    didx = df.columns.get_loc('Delta (Recv-Adv)')
                    ws.conditional_format(1, didx, len(df), didx, {'type': '3_color_scale'})
                    ws.conditional_format(1, didx, len(df), didx, {'type': 'icon_set', 'icon_style': '3_arrows'})
                ws.add_table(0, 0, len(df), len(df.columns)-1, {
                    'style': 'Table Style Light 9', 'columns': [{'header': c} for c in df.columns]
                })
                ws.set_tab_color(theme['tab'])
            summary = _add_summary_sheet({'ALL': all_rows} if mode=='single' else data_per_host, writer, engine, theme)
            state_counts = {}
            if len(all_rows):
                for r in all_rows:
                    st = (r.get('State') or '').title()
                    state_counts[st] = state_counts.get(st, 0) + 1
            _add_dashboard(writer, engine, summary, state_counts, theme)
        else:
            for host, rows_list in data_per_host.items():
                df = pd.DataFrame(rows_list, columns=cols)
                if not df.empty:
                    df['Delta (Recv-Adv)'] = (df['Total Received'] - df['Total Advertised']).astype(int)
                else:
                    df['Delta (Recv-Adv)'] = pd.Series(dtype=int)
                sheet = re.sub(r'[\\/*?:\[\]]', '_', host)[:31] or 'Sheet'
                df.to_excel(writer, index=False, sheet_name=sheet)
                all_rows.extend(rows_list)

                if engine == 'xlsxwriter':
                    wb = writer.book
                    ws = writer.sheets[sheet]
                    header_fmt = wb.add_format({'bold': True, 'bg_color': theme['header_bg'], 'font_color': 'white'})
                    right_num  = wb.add_format({'num_format': '#,##0', 'align': 'right'})
                    left_fmt   = wb.add_format({'align': 'left'})
                    ws.set_row(0, None, header_fmt)
                    ws.freeze_panes(1, 1)
                    _auto_fit_columns(df, ws, wb)
                    numeric_cols = ['IPv4 Received','IPv4 Advertised','IPv6 Received','IPv6 Advertised','Total Received','Total Advertised','Delta (Recv-Adv)']
                    for j, col in enumerate(df.columns):
                        ws.set_column(j, j, None, right_num if col in numeric_cols else left_fmt)
                    if 'State' in df.columns:
                        sidx = df.columns.get_loc('State')
                        ws.conditional_format(1, sidx, len(df), sidx, {
                            'type': 'formula', 'criteria': '=UPPER($%s2)="ESTABLISHED"' % chr(ord('A')+sidx),
                            'format': wb.add_format({'font_color': '#006100', 'bg_color': '#C6EFCE'})
                        })
                        ws.conditional_format(1, sidx, len(df), sidx, {
                            'type': 'formula', 'criteria': '=OR(UPPER($%s2)="IDLE",UPPER($%s2)="ACTIVE",UPPER($%s2)="CONNECT")' % (chr(ord('A')+sidx),chr(ord('A')+sidx),chr(ord('A')+sidx)),
                            'format': wb.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE'})
                        })
                    for col_name in ['Total Received','Total Advertised']:
                        if col_name in df.columns:
                            cidx = df.columns.get_loc(col_name)
                            ws.conditional_format(1, cidx, len(df), cidx, {'type': 'data_bar', 'bar_color': theme['bar']})
                    if 'Delta (Recv-Adv)' in df.columns:
                        didx = df.columns.get_loc('Delta (Recv-Adv)')
                        ws.conditional_format(1, didx, len(df), didx, {'type': '3_color_scale'})
                        ws.conditional_format(1, didx, len(df), didx, {'type': 'icon_set', 'icon_style': '3_arrows'})
                    ws.add_table(0, 0, len(df), len(df.columns)-1, {
                        'style': 'Table Style Light 9', 'columns': [{'header': c} for c in df.columns]
                    })
                    ws.set_tab_color(theme['tab'])

            summary = _add_summary_sheet(data_per_host, writer, engine, theme)
            state_counts = {}
            for r in all_rows:
                st = (r.get('State') or '').title()
                state_counts[st] = state_counts.get(st, 0) + 1
            _add_dashboard(writer, engine, summary, state_counts, theme)


def find_latest_logs_dir(base_dir):
    logs_root = os.path.join(base_dir, 'logs')
    if not os.path.isdir(logs_root):
        return ''
    candidates = []
    for name in os.listdir(logs_root):
        p = os.path.join(logs_root, name)
        if os.path.isdir(p) and len(name) == 8 and name.isdigit():
            try:
                _ = datetime.strptime(name, '%Y%m%d')
                candidates.append((name, p))
            except Exception:
                pass
    if not candidates:
        return ''
    candidates.sort(key=lambda t: t[0], reverse=True)
    return candidates[0][1]


def parse_perhost_dir(logs_dir):
    result = {}
    for name in sorted(os.listdir(logs_dir)):
        if not name.lower().endswith('.txt'):
            continue
        path = os.path.join(logs_dir, name)
        try:
            txt = open(path, 'r', encoding='utf-8', errors='ignore').read()
        except Exception:
            continue
        bgp_text = extract_cmd_block(txt, 'show bgp neighbor')
        if not bgp_text:
            bgp_text = txt
        host_display = os.path.splitext(name)[0]
        parsed = parse_junos_show_bgp_neighbor(bgp_text, host_display)
        result[host_display] = parsed
    return result


def parse_whole_log_as_single(log_text: str):
    """Fallback keras: jika tidak ada header HOST dan tidak ada folder logs,
    tetap coba ekstrak blok 'show bgp neighbor' dari log gabungan dan parse sebagai host 'LOG'."""
    bgp_text = extract_cmd_block(log_text, 'show bgp neighbor')
    if not bgp_text:
        # jika tidak ada marker >>>, pakai seluruh teks
        bgp_text = log_text
    rows = parse_junos_show_bgp_neighbor(bgp_text, 'LOG')
    return {'LOG': rows} if rows else {}


def main():
    ap = argparse.ArgumentParser(description='Parse Junos show bgp neighbor ke Excel (v2.8, fix HOST header & fallback)')
    ap.add_argument('--log', required=True, help='Path file log gabungan')
    ap.add_argument('--hosts', required=False, help='(opsional) Path file host list')
    ap.add_argument('--out', required=True, help='Path output Excel (.xlsx)')
    ap.add_argument('--mode', choices=['per-host', 'single'], default='per-host', help='per-host: sheet per host; single: satu sheet')
    ap.add_argument('--engine', choices=['xlsxwriter','openpyxl','auto'], default='xlsxwriter', help='Prefer engine Excel (xlsxwriter direkomendasikan)')
    args = ap.parse_args()

    if not os.path.exists(args.log):
        sys.stderr.write('Log tidak ditemukan: %s\n' % args.log)
        sys.exit(2)

    base_dir = os.path.dirname(args.log)
    with open(args.log, 'r', encoding='utf-8', errors='ignore') as f:
        log_text = f.read()

    sections = split_by_hosts_via_host_header(log_text)
    data_per_host = {}
    if sections:
        for host, host_text in sections.items():
            bgp_text = extract_cmd_block(host_text, 'show bgp neighbor')
            if not bgp_text:
                bgp_text = host_text
            parsed = parse_junos_show_bgp_neighbor(bgp_text, host)
            data_per_host[host] = parsed
    else:
        logs_dir = find_latest_logs_dir(base_dir)
        if logs_dir:
            data_per_host = parse_perhost_dir(logs_dir)
        if not data_per_host:
            # fallback keras: parse seluruh log sebagai satu host
            data_per_host = parse_whole_log_as_single(log_text)

    # Validasi akhirnya: ada baris data tidak?
    total_rows = sum(len(v) for v in data_per_host.values())
    if total_rows == 0:
        sys.stderr.write('[WARN] Tidak ada data BGP yang berhasil diparse. Cek apakah log berisi output "show bgp neighbor".\n')
        sys.exit(3)

    out_dir = os.path.dirname(args.out)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir)

    engine_pref = 'xlsxwriter' if args.engine in ('xlsxwriter','auto') else args.engine
    to_excel(data_per_host, args.out, mode=args.mode, engine_pref=engine_pref)
    print('Selesai menulis: %s' % args.out)


if __name__ == '__main__':
    main()
