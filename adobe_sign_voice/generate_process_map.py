#!/usr/bin/env python3
"""
generate_process_map.py
========================
Generates a draw.io XML process map for the PSTN Migration Adobe Sign Automation.

Run:    python generate_process_map.py
Output: process_map.drawio

Import into Lucidchart:
    File  >  Import  >  Diagrams.net (draw.io)
"""

from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).parent / "process_map.drawio"

# ── Shape styles ───────────────────────────────────────────────────────────────
TERM  = ("ellipse;whiteSpace=wrap;html=1;"
         "fillColor=#d5e8d4;strokeColor=#82b366;fontStyle=1;fontSize=10;")
PROC  = ("rounded=1;whiteSpace=wrap;html=1;"
         "fillColor=#dae8fc;strokeColor=#6c8ebf;fontSize=10;")
DEC   = ("rhombus;whiteSpace=wrap;html=1;"
         "fillColor=#fff2cc;strokeColor=#d6b656;fontSize=10;verticalAlign=middle;")
DATA  = ("shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;"
         "fillColor=#f8cecc;strokeColor=#b85450;fontSize=10;")
EXT   = ("rounded=1;whiteSpace=wrap;html=1;"
         "fillColor=#f5f5f5;strokeColor=#888888;fontColor=#333333;fontSize=10;dashed=1;")
EDGE  = ("edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;"
         "jettySize=auto;fontSize=9;")
EDGE_EXIT_B = ("edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;"
               "jettySize=auto;fontSize=9;"
               "exitX=0.5;exitY=1;exitDx=0;exitDy=0;")
EDGE_EXIT_R = ("edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;"
               "jettySize=auto;fontSize=9;"
               "exitX=1;exitY=0.5;exitDx=0;exitDy=0;")

def _hdr(color, stroke):
    return (f"rounded=1;whiteSpace=wrap;html=1;"
            f"fillColor={color};strokeColor={stroke};"
            f"fontStyle=1;fontSize=13;")

def _band(color, stroke):
    return (f"rounded=0;whiteSpace=wrap;html=1;"
            f"fillColor={color};strokeColor={stroke};opacity=25;")

HDR = {
    "setup":    _hdr("#fff2cc", "#d6b656"),
    "orch":     _hdr("#f5f5f5", "#888888"),
    "send":     _hdr("#dae8fc", "#6c8ebf"),
    "monitor":  _hdr("#ffe6cc", "#d79b00"),
    "response": _hdr("#e1d5e7", "#9673a6"),
    "reminder": _hdr("#d5e8d4", "#82b366"),
}
BAND = {
    "setup":    _band("#fffde7", "#f5c400"),
    "send":     _band("#e3f2fd", "#1565c0"),
    "monitor":  _band("#fff3e0", "#e65100"),
    "response": _band("#f3e5f5", "#6a1b9a"),
    "reminder": _band("#e8f5e9", "#1b5e20"),
}

# ── ID counter ─────────────────────────────────────────────────────────────────
_ctr = [1]

def nid():
    _ctr[0] += 1
    return str(_ctr[0])


# ── Cell builders ──────────────────────────────────────────────────────────────
cells = []


def V(label, style, x, y, w, h):
    i = nid()
    cells.append(
        f'<mxCell id="{i}" value="{escape(label)}" style="{style}" '
        f'vertex="1" parent="1">'
        f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>'
        f'</mxCell>'
    )
    return i


def E(label, src, tgt, style=EDGE):
    i = nid()
    cells.append(
        f'<mxCell id="{i}" value="{escape(label)}" style="{style}" '
        f'edge="1" source="{src}" target="{tgt}" parent="1">'
        f'<mxGeometry relative="1" as="geometry"/>'
        f'</mxCell>'
    )
    return i


# ══════════════════════════════════════════════════════════════════════════════
# LAYOUT CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
# Canvas: 2600 wide
# Section horizontal content starts at x=360 (section labels occupy 0-340)
CX   = 360    # content start x
STEP = 205    # standard horizontal step (170 node + 35 gap)
GAP  = 35     # gap between nodes

PW, PH     = 170, 55   # process node
TW, TH     = 120, 50   # terminal node
DW, DH     = 130, 90   # decision node
DAW, DAH   = 170, 55   # data / parallelogram node
EW, EH     = 190, 55   # external node

# Section Y positions (top of background band)
SY = {
    "setup":    20,
    "orch":     250,
    "send":     465,
    "monitor":  860,
    "response": 1305,
    "reminder": 1790,
}
SH = {
    "setup":    200,
    "send":     360,
    "monitor":  415,
    "response": 455,
    "reminder": 415,
}

# ── Background bands ───────────────────────────────────────────────────────────
for k in ("setup", "send", "monitor", "response", "reminder"):
    V("", BAND[k], 15, SY[k], 2570, SH[k])

# ── Section header labels ──────────────────────────────────────────────────────
V("1.  ONE-TIME SETUP\n(run once before first use)",
  HDR["setup"],    15, SY["setup"],    335, 45)
V("2.  ORCHESTRATOR\n(Windows Task Scheduler)",
  HDR["orch"],     15, SY["orch"],     335, 45)
V("3.  BULK SEND\n(orchestrator.py --send)",
  HDR["send"],     15, SY["send"],     335, 45)
V("4.  AGREEMENT MONITOR\n(orchestrator.py --monitor)",
  HDR["monitor"],  15, SY["monitor"],  335, 45)
V("5.  RESPONSE PROCESSOR\n(triggered by Monitor on COMPLETED)",
  HDR["response"], 15, SY["response"], 335, 45)
V("6.  REMINDER SENDER\n(orchestrator.py --reminders)",
  HDR["reminder"], 15, SY["reminder"], 335, 45)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — ONE-TIME SETUP
# ══════════════════════════════════════════════════════════════════════════════
R1 = SY["setup"] + 85   # node row y

s1_start   = V("START\n(First-time setup)",           TERM,  CX,           R1,       TW, TH)
s1_oauth   = V("oauth_setup.py\n(opens browser)",     PROC,  CX+TW+GAP,    R1,       PW, PH)
s1_browser = V("Browser: Authorise\nAdobe Sign app",  EXT,   CX+TW+GAP+PW+GAP, R1,   EW, EH)
s1_tokens  = V("Access + refresh\ntokens received",   PROC,  CX+TW+GAP+PW+GAP+EW+GAP, R1, PW, PH)
s1_save    = V("refresh_token written\nto .env",      DATA,  CX+TW+GAP+PW+GAP+EW+GAP+PW+GAP, R1, DAW, DAH)
s1_list    = V("list_library_docs.py\n(lists templates)", PROC,
               CX+TW+GAP+PW+GAP+EW+GAP+PW+GAP+DAW+GAP, R1, PW, PH)
s1_docid   = V("Copy Library Doc ID\ninto ADOBE_SIGN_LIBRARY_DOC_ID in .env",
               DATA, CX+TW+GAP+PW+GAP+EW+GAP+PW+GAP+DAW+GAP+PW+GAP, R1, DAW+30, DAH)
s1_done    = V("SETUP\nCOMPLETE",                     TERM,
               CX+TW+GAP+PW+GAP+EW+GAP+PW+GAP+DAW+GAP+PW+GAP+DAW+30+GAP, R1, TW, TH)

E("", s1_start,   s1_oauth)
E("", s1_oauth,   s1_browser)
E("", s1_browser, s1_tokens)
E("", s1_tokens,  s1_save)
E("", s1_save,    s1_list)
E("", s1_list,    s1_docid)
E("", s1_docid,   s1_done)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════════════════
R2   = SY["orch"] + 70
R2b  = R2 + 75   # second row (mode labels)

s2_sched = V("Windows Task Scheduler\n(runs daily / on demand)", EXT,  CX,             R2, EW,    EH)
s2_orch  = V("orchestrator.py",                                   PROC, CX+EW+GAP,      R2, PW,    PH)
s2_fork  = V("Dispatch\nmode?",                                   DEC,  CX+EW+GAP+PW+GAP, R2-(DH-PH)//2, DW, DH)

# Mode label boxes (no edges between them — just informational)
fx = CX + EW + GAP + PW + GAP + DW + 50
s2_send_lbl  = V("--send\ncalls bulk_sender.py",            PROC, fx,       R2b-20, 180, 45)
s2_mon_lbl   = V("--monitor (default)\ncalls agreement_monitor.py", PROC, fx, R2b+40, 230, 45)
s2_rem_lbl   = V("--reminders (default)\ncalls reminder_sender.py", PROC, fx, R2b+100, 230, 45)
s2_stat_lbl  = V("--status\nprints state summary to stdout", EXT,  fx,       R2b+160, 230, 45)

E("--send",           s2_fork, s2_send_lbl,  EDGE)
E("--monitor",        s2_fork, s2_mon_lbl,   EDGE)
E("--reminders",      s2_fork, s2_rem_lbl,   EDGE)
E("--status",         s2_fork, s2_stat_lbl,  EDGE)
E("",                 s2_sched, s2_orch)
E("",                 s2_orch,  s2_fork)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — BULK SEND
# ══════════════════════════════════════════════════════════════════════════════
R3a = SY["send"] + 75    # row 1
R3b = SY["send"] + 210   # row 2

# Row 1
s3_bulk    = V("bulk_sender.py\nrun_bulk_send()",             PROC, CX,           R3a, PW,    PH)
s3_load    = V("sharepoint_reader\n.load_master_data()",      PROC, CX+STEP,      R3a, PW,    PH)
s3_filter  = V("Filter: skip rows\nwhere date_sent filled",   PROC, CX+STEP*2,    R3a, PW,    PH)
s3_dec_any = V("Any unsent\npartners?",                       DEC,  CX+STEP*3,    R3a-(DH-PH)//2, DW, DH)
s3_skip    = V("Nothing to send.\nExit.",                     TERM, CX+STEP*3+DW+GAP*2, R3a, TW, TH)

# Row 2 (Yes branch — loop body)
s3_each    = V("For each unsent\npartner",                    PROC, CX,           R3b, PW,    PH)
s3_merge   = V("build_merge_fields()\n(partner data → Adobe fields)", PROC, CX+STEP, R3b, PW, PH)
s3_create  = V("client.create_agreement()\n(Adobe Sign REST API v6)", PROC, CX+STEP*2, R3b, PW, PH)
s3_state   = V("Append record to\nagreement_state.json",      DATA, CX+STEP*3,    R3b, DAW,   DAH)
s3_excel   = V("try_update_excel_after_send()\n(write date_sent + agreement_id)", PROC,
               CX+STEP*3+DAW+GAP, R3b, PW+20, PH)
s3_delay   = V("Sleep\nSEND_DELAY_SECONDS",                   PROC, CX+STEP*3+DAW+GAP+PW+20+GAP, R3b, 140, PH)
s3_end     = V("SEND\nCOMPLETE",                              TERM, CX+STEP*3+DAW+GAP+PW+20+GAP+140+GAP, R3b, TW, TH)

E("",    s3_bulk,    s3_load)
E("",    s3_load,    s3_filter)
E("",    s3_filter,  s3_dec_any)
E("No",  s3_dec_any, s3_skip,   EDGE_EXIT_R)
E("Yes", s3_dec_any, s3_each,   EDGE_EXIT_B)
E("",    s3_each,    s3_merge)
E("",    s3_merge,   s3_create)
E("",    s3_create,  s3_state)
E("",    s3_state,   s3_excel)
E("",    s3_excel,   s3_delay)
E("",    s3_delay,   s3_end)

# Arrow: orchestrator --send → bulk_sender
E("--send", s2_send_lbl, s3_bulk, EDGE)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — AGREEMENT MONITOR
# ══════════════════════════════════════════════════════════════════════════════
R4a = SY["monitor"] + 75
R4b = SY["monitor"] + 215

# Row 1
s4_mon     = V("agreement_monitor.py\nrun_monitor()",          PROC, CX,           R4a, PW,    PH)
s4_load    = V("Load\nagreement_state.json",                   DATA, CX+STEP,      R4a, DAW,   DAH)
s4_filter  = V("Filter records where\nstatus = OUT_FOR_SIGNATURE", PROC, CX+STEP+DAW+GAP, R4a, PW, PH)
s4_dec_any = V("Any pending\nagreements?",                     DEC,  CX+STEP+DAW+GAP+PW+GAP, R4a-(DH-PH)//2, DW, DH)
s4_none    = V("Nothing to check.\nExit.",                     TERM, CX+STEP+DAW+GAP+PW+GAP+DW+GAP*2, R4a, TW, TH)

# Row 2
s4_each    = V("For each pending\nagreement",                  PROC, CX,           R4b, PW,    PH)
s4_get     = V("client.get_agreement()\n(poll status from Adobe Sign)", PROC, CX+STEP, R4b, PW, PH)
s4_dec_st  = V("Agreement\nstatus?",                           DEC,  CX+STEP*2,    R4b-(DH-PH)//2, DW, DH)

# Three branches
bx = CX + STEP*2 + DW + GAP
s4_comp    = V("SIGNED / COMPLETED\n→ trigger response_processor",
               PROC, bx,           R4b-80, PW+10, PH)
s4_canc    = V("CANCELLED / DECLINED\nUpdate state. Log for manual review.",
               PROC, bx,           R4b+5,  PW+10, PH)
s4_pend    = V("STILL PENDING\nNo action — checked next run.",
               PROC, bx,           R4b+90, PW+10, PH)

E("",               s4_mon,     s4_load)
E("",               s4_load,    s4_filter)
E("",               s4_filter,  s4_dec_any)
E("No",             s4_dec_any, s4_none,   EDGE_EXIT_R)
E("Yes",            s4_dec_any, s4_each,   EDGE_EXIT_B)
E("",               s4_each,    s4_get)
E("",               s4_get,     s4_dec_st)
E("COMPLETED",      s4_dec_st,  s4_comp,   EDGE)
E("CANCELLED",      s4_dec_st,  s4_canc,   EDGE)
E("OTHER",          s4_dec_st,  s4_pend,   EDGE)

# Arrow: orchestrator --monitor → agreement_monitor
E("--monitor", s2_mon_lbl, s4_mon, EDGE)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — RESPONSE PROCESSOR
# ══════════════════════════════════════════════════════════════════════════════
R5a = SY["response"] + 75
R5b = SY["response"] + 215
R5c = SY["response"] + 355

# Row 1
s5_entry   = V("process_signed_agreement()\nagreement_id, partner_name",  PROC, CX,        R5a, PW+10, PH)
s5_form    = V("client.get_form_data()\n(download CSV form fields)",       PROC, CX+STEP+10, R5a, PW,    PH)
s5_savejson= V("Save form data →\noutput/{name}_form_data_{ts}.json",     DATA, CX+STEP*2+10, R5a, DAW+10, DAH)
s5_docs    = V("client.get_documents()\n(list agreement documents)",       PROC, CX+STEP*3+20, R5a, PW,    PH)
s5_xls_dl  = V("Download Excel\nattachments from signer",                 DATA, CX+STEP*4+20, R5a, DAW,   DAH)
s5_csv     = V("Parse Excel → DataFrame\nSave as CSV to output/",         DATA, CX+STEP*5+20, R5a, DAW+10, DAH)

# Row 2
s5_pdf_dl  = V("client.download_combined\n_document()\n(full signed PDF)", PROC, CX,        R5b, PW+10, PH)
s5_savepdf = V("Save signed PDF →\noutput/{name}_signed_{ts}.pdf",        DATA, CX+STEP+10, R5b, DAW,   DAH)
s5_xl_up   = V("try_update_excel_after\n_completion()\nWrite date_received", PROC,
               CX+STEP+10+DAW+GAP, R5b, PW+10, PH)
s5_state   = V("Mark processed=True\nin agreement_state.json",             DATA, CX+STEP+10+DAW+GAP+PW+10+GAP, R5b, DAW, DAH)
s5_done    = V("PROCESSING\nCOMPLETE",                                     TERM, CX+STEP+10+DAW+GAP+PW+10+GAP+DAW+GAP, R5b, TW, TH)

E("", s5_entry,  s5_form)
E("", s5_form,   s5_savejson)
E("", s5_savejson, s5_docs)
E("", s5_docs,   s5_xls_dl)
E("", s5_xls_dl, s5_csv)
# Row 1 wraps to row 2
E("", s5_csv,    s5_pdf_dl,  EDGE)
E("", s5_pdf_dl, s5_savepdf)
E("", s5_savepdf, s5_xl_up)
E("", s5_xl_up,  s5_state)
E("", s5_state,  s5_done)

# Arrow: COMPLETED branch → response_processor
E("process_signed_agreement()", s4_comp, s5_entry, EDGE)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — REMINDER SENDER
# ══════════════════════════════════════════════════════════════════════════════
R6a = SY["reminder"] + 75
R6b = SY["reminder"] + 210

# Row 1
s6_rem     = V("reminder_sender.py\nrun_reminders()",                    PROC, CX,           R6a, PW,    PH)
s6_load    = V("Load\nagreement_state.json",                             DATA, CX+STEP,      R6a, DAW,   DAH)
s6_filter  = V("Filter: OUT_FOR_SIGNATURE\nrecords only",                PROC, CX+STEP+DAW+GAP, R6a, PW, PH)
s6_token   = V("Acquire Microsoft Graph\nOAuth token (MSAL)",            PROC, CX+STEP+DAW+GAP+PW+GAP, R6a, PW+10, PH)
s6_dec_any = V("Any pending\nagreements?",                               DEC,  CX+STEP+DAW+GAP+PW+GAP+PW+10+GAP, R6a-(DH-PH)//2, DW, DH)
s6_none    = V("Nothing due.\nExit.",                                    TERM, CX+STEP+DAW+GAP+PW+GAP+PW+10+GAP+DW+GAP*2, R6a, TW, TH)

# Row 2 — per-agreement reminder logic (3 day thresholds)
s6_each    = V("For each pending\nagreement",                            PROC, CX,           R6b, PW,    PH)
s6_days    = V("Days since\nsent_date?",                                 DEC,  CX+STEP,      R6b-(DH-PH)//2, DW, DH)

dx = CX + STEP + DW + GAP
row6_gap = 120

# Day-7 path
s6_7_get   = V("client.get_signing_url()\n(fetch live signing link)",    PROC, dx,           R6b - row6_gap, PW,    PH)
s6_7_send  = V("Send day-7 chaser\nemail via Microsoft Graph",           PROC, dx+STEP,      R6b - row6_gap, PW,    PH)
s6_7_mark  = V("Set reminder_7_sent\n= True  in state.json",            DATA, dx+STEP*2,    R6b - row6_gap, DAW,   DAH)

# Day-14 path
s6_14_get  = V("client.get_signing_url()\n(fetch live signing link)",    PROC, dx,           R6b,            PW,    PH)
s6_14_send = V("Send day-14 chaser\nemail via Microsoft Graph",          PROC, dx+STEP,      R6b,            PW,    PH)
s6_14_mark = V("Set reminder_14_sent\n= True  in state.json",           DATA, dx+STEP*2,    R6b,            DAW,   DAH)

# Day-30 path
s6_30_get  = V("client.get_signing_url()\n(fetch live signing link)",    PROC, dx,           R6b + row6_gap, PW,    PH)
s6_30_send = V("Send day-30 chaser\nemail via Microsoft Graph",          PROC, dx+STEP,      R6b + row6_gap, PW,    PH)
s6_30_mark = V("Set reminder_30_sent\n= True  in state.json",           DATA, dx+STEP*2,    R6b + row6_gap, DAW,   DAH)

s6_save    = V("Save\nagreement_state.json",                             DATA, dx+STEP*3,    R6b,            DAW,   DAH)
s6_done    = V("REMINDERS\nCOMPLETE",                                    TERM, dx+STEP*3+DAW+GAP, R6b,      TW, TH)

E("",      s6_rem,     s6_load)
E("",      s6_load,    s6_filter)
E("",      s6_filter,  s6_token)
E("",      s6_token,   s6_dec_any)
E("No",    s6_dec_any, s6_none,    EDGE_EXIT_R)
E("Yes",   s6_dec_any, s6_each,    EDGE_EXIT_B)
E("",      s6_each,    s6_days)
E(">= 7d", s6_days,    s6_7_get,   EDGE)
E(">= 14d",s6_days,    s6_14_get,  EDGE)
E(">= 30d",s6_days,    s6_30_get,  EDGE)
E("",      s6_7_get,   s6_7_send)
E("",      s6_7_send,  s6_7_mark)
E("",      s6_14_get,  s6_14_send)
E("",      s6_14_send, s6_14_mark)
E("",      s6_30_get,  s6_30_send)
E("",      s6_30_send, s6_30_mark)
E("",      s6_7_mark,  s6_save,    EDGE)
E("",      s6_14_mark, s6_save,    EDGE)
E("",      s6_30_mark, s6_save,    EDGE)
E("",      s6_save,    s6_done)

# Arrow: orchestrator --reminders → reminder_sender
E("--reminders", s2_rem_lbl, s6_rem, EDGE)


# ══════════════════════════════════════════════════════════════════════════════
# BUILD XML
# ══════════════════════════════════════════════════════════════════════════════
canvas_h = SY["reminder"] + SH["reminder"] + 60

xml_lines = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    f'<mxGraphModel compressed="false" '
    f'dx="1422" dy="762" grid="1" gridSize="10" '
    f'guides="1" tooltips="1" connect="1" arrows="1" fold="1" '
    f'page="1" pageScale="1" pageWidth="2600" pageHeight="{canvas_h}" '
    f'math="0" shadow="0">',
    '  <root>',
    '    <mxCell id="0"/>',
    '    <mxCell id="1" parent="0"/>',
]
xml_lines += [f'    {c}' for c in cells]
xml_lines += ['  </root>', '</mxGraphModel>']

xml = "\n".join(xml_lines)
OUT.write_text(xml, encoding="utf-8")

print(f"Process map written to: {OUT}")
print()
print("Import into Lucidchart:")
print("  1. Open Lucidchart")
print("  2. File > Import > Diagrams.net (draw.io)")
print(f"  3. Select: {OUT}")
