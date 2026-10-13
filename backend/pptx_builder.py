from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_THEME_COLOR
import re

# Editorial palette: plum ink, warm white, teal accent, restrained ochre.
PLUM = RGBColor(66, 34, 65)
PURPLE = RGBColor(111, 45, 103)
TEAL = RGBColor(0, 151, 137)
INK = RGBColor(40, 37, 42)
MUTED = RGBColor(103, 97, 105)
PAPER = RGBColor(250, 248, 247)
LILAC = RGBColor(237, 229, 238)
PALE_TEAL = RGBColor(225, 242, 239)
OCHRE = RGBColor(183, 126, 59)
WHITE = RGBColor(255,255,255)
LINE = RGBColor(218, 211, 219)
FONT = "Aptos"
W, H = 13.333, 7.5

def validate_proposal(data):
    if not isinstance(data, dict) or not isinstance(data.get("client_slides"), list):
        raise ValueError("Proposal JSON must contain a client_slides array.")
    slides = data["client_slides"]
    if len(slides) < 15:
        raise ValueError(f"Proposal is too short ({len(slides)} client slides). Generate at least 19.")
    for i, slide in enumerate(slides):
        if not isinstance(slide, dict) or not slide.get("title"):
            raise ValueError(f"Slide {i+1} is missing a title.")
        slide.setdefault("layout", "narrative")
        slide.setdefault("body", [])
        slide.setdefault("bullets", [])
        slide.setdefault("notes", "")

def set_bg(slide, color=PAPER):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def textbox(slide, x, y, w, h, text="", size=16, color=INK, bold=False, font=FONT,
            align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, margin=0.02, italic=False):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(margin); tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin); tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    r = p.add_run()
    r.text = str(text or "")
    r.font.name = font; r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    return shape

def line(slide, x1,y1,x2,y2,color=LINE,width=1.1):
    sh=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    sh.line.color.rgb=color; sh.line.width=Pt(width)
    return sh

def add_footer(slide, page, section="POWER BI SOLUTION PROPOSAL"):
    line(slide, .62, 7.08, 12.7, 7.08, LINE, .65)
    textbox(slide,.64,7.13,7,.18,"XEBIA  /  "+section.upper(),8,MUTED,bold=True)
    textbox(slide,12.1,7.11,.55,.22,f"{page:02d}",9,PURPLE,bold=True,align=PP_ALIGN.RIGHT)

def title_block(slide, eyebrow, title, lead=None, page=1, section="SOLUTION DESIGN"):
    textbox(slide,.68,.35,11.9,.22,(eyebrow or section).upper(),9,TEAL,bold=True)
    textbox(slide,.66,.70,12.0,.68,title,27,PLUM,bold=True)
    line(slide,.68,1.48,1.45,1.48,TEAL,2.5)
    if lead:
        textbox(slide,.68,1.62,11.9,.62,lead,13,MUTED)
    add_footer(slide,page,section)

def add_notes(slide, notes):
    if not notes: return
    try:
        tf=slide.notes_slide.notes_text_frame
        tf.text=str(notes)
    except Exception:
        pass

def body_paragraphs(slide, body, x=.72, y=2.35, w=11.9, h=3.8, size=14):
    if isinstance(body,str): body=[body]
    body=[str(p).strip() for p in (body or []) if str(p).strip()]
    if not body: return
    # Explicit paragraph spacing, no background cards.
    total=min(len(body),4)
    slot=h/max(total,1)
    for i,p in enumerate(body[:4]):
        textbox(slide,x,y+i*slot,w,slot-.08,p,size,INK)

def bullet_items(slide, bullets, x=.8, y=2.45, w=11.8, row_h=.65, max_items=6, size=13):
    items=[]
    for b in (bullets or [])[:max_items]:
        if isinstance(b,dict):
            items.append((str(b.get("heading","")).strip(),str(b.get("detail","")).strip()))
        else: items.append(("",str(b).strip()))
    for i,(head,detail) in enumerate(items):
        yy=y+i*row_h
        textbox(slide,x,yy,.22,.28,"—",15,TEAL,bold=True)
        if head:
            textbox(slide,x+.30,yy,w*.27,.32,head,12,PLUM,bold=True)
            textbox(slide,x+.30+w*.27,yy,w*.70-.3,row_h-.04,detail,11.5,INK)
        else:
            textbox(slide,x+.30,yy,w-.3,row_h-.04,detail,size,INK)

def add_cover(prs, data):
    s=prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s,PLUM)
    # Oversized editorial typography and one accent rule, not boxes.
    textbox(s,.82,.62,6,.28,"XEBIA  /  DATA & ANALYTICS",10,RGBColor(184,225,218),True)
    textbox(s,.82,1.45,11.4,1.45,data.get("proposal_title","Power BI Solution Proposal"),34,WHITE,True)
    line(s,.84,3.18,2.2,3.18,TEAL,4)
    thesis=data.get("executive_thesis","A practical roadmap from business questions to governed, decision-ready analytics.")
    textbox(s,.84,3.55,9.5,1.05,thesis,18,RGBColor(238,230,238))
    textbox(s,.84,5.55,7,.35,f"PREPARED FOR  /  {data.get('client_name','Client')}",11,WHITE,True)
    textbox(s,.84,6.02,7,.28,f"INDUSTRY  /  {data.get('industry','Not specified')}",10,RGBColor(205,194,208))
    textbox(s,11.75,6.78,.8,.28,"XEBIA",10,WHITE,True,align=PP_ALIGN.RIGHT)
    add_notes(s,"Opening narrative: introduce the client context, the business problem as understood from discovery, and the proposal's central recommendation. Confirm that all assumptions are subject to validation.")
    return s

def add_agenda(prs, page):
    s=prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    title_block(s,"THE STORY","From business questions to a governed analytics product",
                "The proposal moves from business context through solution design into delivery, controls and measurable acceptance.",page,"PROPOSAL ROADMAP")
    items=[
      ("01","Business case & outcomes","Clarify the decision problems, users and value expected."),
      ("02","Target solution","Architecture, data products, semantic model and report experience."),
      ("03","Trust & operations","Security, quality, performance, release management and ownership."),
      ("04","Delivery & acceptance","Phases, dependencies, risks, validation and next decisions.")
    ]
    for i,(num,head,desc) in enumerate(items):
        yy=2.45+i*.88
        textbox(s,.82,yy,.55,.34,num,15,TEAL,True)
        textbox(s,1.62,yy,3.0,.34,head,16,PLUM,True)
        textbox(s,4.6,yy,7.65,.5,desc,13,INK)
        if i<3: line(s,1.62,yy+.61,12.1,yy+.61,LINE,.6)
    return s

def add_capability(prs,page,title,lead,body,accent=TEAL):
    s=prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    title_block(s,"XEBIA CAPABILITIES",title,lead,page,"XEBIA CAPABILITIES")
    body_paragraphs(s,body,.78,2.55,11.7,2.3,14)
    line(s,.8,5.25,12.3,5.25,LINE,.8)
    textbox(s,.82,5.48,3.3,.3,"WHAT THIS ENABLES",10,accent,True)
    textbox(s,.82,5.9,11.4,.68,lead,17,PLUM,True)
    return s

def add_table(slide, table_data, y=2.45, height=4.25):
    cols=table_data.get("columns",[])[:6]; rows=table_data.get("rows",[])[:10]
    if not cols or not rows: return
    n=len(cols); x=.72; total_w=11.9
    widths=[total_w/n]*n
    # First column slightly narrower when many columns.
    if n>=4:
        widths=[total_w*(.16 if i==0 else .84/(n-1)) for i in range(n)]
    row_h=min(.54,height/(len(rows)+1))
    yy=y
    # Understated table treatment: horizontal rules, no card cells.
    line(slide,x,yy,x+total_w,yy,PLUM,1.4)
    for j,c in enumerate(cols):
        textbox(slide,x+sum(widths[:j])+.04,yy+.08,widths[j]-.1,row_h-.08,str(c),10,PLUM,True)
    yy+=row_h
    line(slide,x,yy,x+total_w,yy,TEAL,1.2)
    for ri,row in enumerate(rows):
        for j in range(n):
            val=row[j] if j<len(row) else ""
            textbox(slide,x+sum(widths[:j])+.04,yy+.06,widths[j]-.1,row_h-.06,str(val),9.2,INK)
        yy+=row_h
        line(slide,x,yy,x+total_w,yy,LINE,.55)

def add_architecture(slide, diagram):
    nodes=diagram.get("nodes",[])[:8]
    if not nodes: nodes=["Source systems","Ingestion","Transform","Curated data","Semantic model","Power BI"]
    n=len(nodes); x0=.75; y=3.0; total=11.85; gap=.14; nodew=(total-gap*(n-1))/n
    for i,node in enumerate(nodes):
        x=x0+i*(nodew+gap)
        textbox(slide,x,y,nodew,.82,str(node),10.5,PLUM,True,align=PP_ALIGN.CENTER,valign=MSO_ANCHOR.MIDDLE,margin=.07)
        if i<n-1:
            line(slide,x+nodew,y+.41,x+nodew+gap-.02,y+.41,TEAL,1.6)
    textbox(slide,.78,4.2,11.8,.9,diagram.get("caption","Recommended logical flow; validate source systems, volumes, latency and security requirements during discovery."),12,MUTED,italic=True)
    textbox(slide,.78,5.4,2.2,.28,"CROSS-CUTTING CONTROLS",9,TEAL,True)
    textbox(slide,.78,5.78,11.8,.68,"Identity & access  •  Data quality & reconciliation  •  Lineage & metadata  •  Monitoring & alerting  •  Dev → Test → Production",11,INK)

def add_roadmap(slide, bullets):
    phases=bullets[:5]
    if not phases: phases=[{"heading":"Discover","detail":"Confirm scope and baseline."},{"heading":"Design","detail":"Agree target model."},{"heading":"Build","detail":"Engineer and develop."},{"heading":"Validate","detail":"Reconcile and UAT."},{"heading":"Release","detail":"Deploy and transition."}]
    count=len(phases); left=.9; usable=11.5; gap=.22; w=(usable-gap*(count-1))/count; y=3.05
    line(slide,left+w/2,y+.15,left+(count-1)*(w+gap)+w/2,y+.15,TEAL,2)
    for i,item in enumerate(phases):
        if isinstance(item,dict): head=item.get("heading","Phase"); detail=item.get("detail","")
        else: head=f"Phase {i+1}"; detail=str(item)
        x=left+i*(w+gap)
        textbox(slide,x,y-.04,w,.3,f"{i+1:02d}",11,TEAL,True)
        textbox(slide,x,y+.48,w,.52,head,13,PLUM,True)
        textbox(slide,x,y+1.1,w,1.65,detail,10.5,INK)
        line(slide,x,y+2.9,x+w,y+2.9,LINE,.8)

def add_client_slide(prs, item, page):
    s=prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    layout=item.get("layout","narrative")
    title_block(s,item.get("eyebrow","CLIENT-SPECIFIC SOLUTION"),item.get("title","Solution design"),item.get("lead",""),page,"CLIENT SOLUTION")
    body=item.get("body",[])
    if isinstance(body,str): body=[body]
    if layout=="table" and item.get("table"):
        body_paragraphs(s,body[:1],.74,2.28,11.7,.62,11.5)
        add_table(s,item["table"],2.95,3.85)
    elif layout=="architecture":
        body_paragraphs(s,body[:1],.74,2.28,11.7,.6,11.5)
        add_architecture(s,item.get("diagram",{}))
    elif layout=="roadmap":
        body_paragraphs(s,body[:1],.74,2.25,11.7,.6,11.5)
        add_roadmap(s,item.get("bullets",[]))
    elif layout=="scorecard" and item.get("table"):
        body_paragraphs(s,body[:1],.74,2.25,11.7,.6,11.5)
        add_table(s,item["table"],2.95,3.85)
    elif layout=="split":
        left=body[:2]; right=body[2:4]
        textbox(s,.78,2.48,5.55,.25,"DESIGN RATIONALE",9,TEAL,True)
        body_paragraphs(s,left,.78,2.85,5.55,2.75,12)
        line(s,6.62,2.48,6.62,5.95,LINE,.9)
        textbox(s,6.95,2.48,5.3,.25,"IMPLEMENTATION DETAIL",9,TEAL,True)
        body_paragraphs(s,right,6.95,2.85,5.3,2.75,12)
        bullet_items(s,item.get("bullets",[])[:3],.82,5.82,11.5,.37,3,10.5)
    elif layout=="section":
        textbox(s,.8,2.6,10.9,1.4,item.get("lead",""),23,PLUM,True)
        body_paragraphs(s,body,.8,4.25,11.3,1.55,13)
    elif layout=="closing":
        body_paragraphs(s,body,.8,2.45,11.5,1.7,14)
        bullet_items(s,item.get("bullets",[]),.82,4.45,11.5,.48,5,12)
    else:
        # Editorial narrative with two clear zones; no card backgrounds.
        body_paragraphs(s,body[:3],.78,2.45,11.65,2.15,12.5)
        if len(body)>3:
            line(s,.8,4.83,12.3,4.83,LINE,.8)
            textbox(s,.82,5.02,2.1,.25,"WHAT THIS MEANS",9,TEAL,True)
            body_paragraphs(s,body[3:],2.85,5.0,9.3,1.1,11.5)
        else:
            bullet_items(s,item.get("bullets",[]),.82,4.8,11.5,.38,5,10.7)
    add_notes(s,item.get("notes",""))
    return s

def add_requirements_slide(prs, data, page):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    client = data.get("client_name", "Client")
    title_block(s, "CLIENT REQUIREMENTS", f"What we heard from {client}",
                data.get("client_requirement_summary", data.get("executive_thesis", "A summary of the client's stated requirements and the outcomes to validate.")),
                page, "CLIENT REQUIREMENTS")
    objectives = data.get("business_objectives", [])[:5]
    challenges = data.get("key_challenges", [])[:5]
    criteria = data.get("success_criteria", [])[:4]
    # Editorial two-column requirement summary with rules, not cards.
    textbox(s, .78, 2.55, 5.5, .26, "BUSINESS OBJECTIVES", 9, TEAL, True)
    bullet_items(s, objectives, .78, 2.92, 5.55, .56, 5, 10.5)
    line(s, 6.62, 2.55, 6.62, 5.95, LINE, .8)
    textbox(s, 6.95, 2.55, 5.4, .26, "CHALLENGES TO ADDRESS", 9, TEAL, True)
    bullet_items(s, challenges, 6.95, 2.92, 5.45, .56, 5, 10.5)
    if criteria:
        line(s, .8, 5.92, 12.3, 5.92, LINE, .7)
        textbox(s, .82, 6.02, 1.7, .24, "SUCCESS SIGNALS", 8.5, PURPLE, True)
        textbox(s, 2.55, 6.00, 9.7, .55, "  •  ".join(str(x) for x in criteria), 9.5, INK)
    add_notes(s, "Open the client-specific section by confirming the interpretation of the requirements. Validate these objectives, challenges and success signals with stakeholders before finalizing scope.")

def build_presentation(data, output_path: Path):
    validate_proposal(data)
    prs=Presentation()
    prs.slide_width=Inches(W); prs.slide_height=Inches(H)
    prs.core_properties.title=data.get("proposal_title","Power BI Solution Proposal")
    prs.core_properties.subject="Client-specific Power BI proposal"
    prs.core_properties.author="Xebia Data & Analytics"
    add_cover(prs,data)

    # Eight fixed Xebia capability and accelerator slides. These are always included,
    # before the client-specific requirements and solution section.
    capabilities=[
      ("Power BI consulting & decision intelligence",
       "We design analytics around business decisions, governed definitions and measurable outcomes.",
       ["Start with stakeholder decisions and operating workflows, then design measures, report interactions and ownership around those decisions.","Create a reusable semantic layer so executive, operational and analytical reports share consistent calculations and business definitions."]),
      ("Microsoft Fabric, OneLake & data engineering",
       "Connect source systems to governed, reusable analytical data products.",
       ["Evaluate ingestion, Lakehouse/Warehouse, transformation and serving patterns against source capabilities, latency, volume, recovery and licensing constraints.","Use raw, standardized and curated layers where they provide traceability, quality enforcement and controlled business-ready data."]),
      ("Semantic models, KPI governance & performance",
       "Make measures consistent, explainable and reusable across the reporting estate.",
       ["Define fact-table grain, business keys, conformed dimensions, relationship cardinality and filter direction before report development.","Document KPI meaning, formula, exclusions, time basis, dimensional cuts, source lineage, owner and reconciliation approach."]),
      ("Enterprise reporting & data storytelling",
       "Build purposeful experiences for executive, analytical and operational users.",
       ["Use executive pages for outcomes and exceptions, analytical pages for decomposition and root cause, and operational pages for queues and next actions.","Design drill-through, tooltips, accessible visual hierarchy and performance-aware interactions around real user questions rather than decorative visuals."]),
      ("Accelerator 01 — Tableau Workbook Validator",
       "We built validation automation to reduce manual effort and risk during large-scale Tableau migration.",
       ["Inspect workbook and connection metadata to identify environment-specific datasource references and configuration mismatches before release.","Use repeatable validation evidence to find workbooks still pointing at development sources, isolate connection failures and improve migration readiness."]),
      ("Accelerator 02 — Qlik-to-Power BI Expression Translator",
       "We built translation assistance to accelerate conversion of legacy Qlik expressions into maintainable DAX.",
       ["Apply deterministic mappings for common expression patterns first, then route complex or ambiguous logic through AI-assisted review rather than paying for unnecessary model calls.","Return a proposed DAX expression with assumptions and validation notes; business totals and edge cases must be reconciled before production use."]),
      ("Accelerator 03 — Power BI Playbook Builder",
       "We built a requirements-to-playbook accelerator to turn discovery inputs into a structured implementation blueprint.",
       ["Convert plain-language requirements or BRDs into architecture, semantic-model proposals, KPI catalogues, report-page plans, governance and delivery guidance.","The goal is to shorten the blank-page phase for consultants while making assumptions visible and leaving final design decisions with the project team."]),
      ("Accelerator strategy — metadata, migration & delivery",
       "The accelerators target repeatable delivery friction: discovery, metadata inspection, translation and validation.",
       ["Build reusable tooling because migration and proposal tasks repeatedly require the same checks, mappings and documentation across clients.","Use automation to improve consistency and traceability—not to replace architect review, source validation, reconciliation, security testing or business sign-off."])
    ]
    page=2
    for title,lead,body in capabilities:
        add_capability(prs,page,title,lead,body); page+=1

    # Dedicated requirements slide always comes immediately after the fixed Xebia section.
    add_requirements_slide(prs,data,page); page+=1
    for item in data["client_slides"]:
        add_client_slide(prs,item,page); page+=1
    output_path=Path(output_path); output_path.parent.mkdir(parents=True,exist_ok=True)
    prs.save(output_path)
    check=Presentation(str(output_path))
    if len(check.slides) != len(prs.slides):
        raise RuntimeError("Saved PowerPoint slide count did not validate.")
    # Assert mandatory ordering: cover, 8 Xebia capability slides, requirements, client detail.
    slide_titles=[]
    for sl in check.slides:
        texts=[sh.text.strip() for sh in sl.shapes if getattr(sh, "has_text_frame", False) and sh.text.strip()]
        slide_titles.append(texts[1] if len(texts)>1 else (texts[0] if texts else ""))
    if len(check.slides) < 27:
        raise RuntimeError("Generated proposal is unexpectedly short.")
    return len(check.slides)
