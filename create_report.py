#!/usr/bin/env python3
"""
Generate Amul Internship Report as a .docx file using only standard library.
A .docx file is a ZIP archive containing XML files following the Office Open XML standard.
"""

import zipfile
import os

def create_docx(filename):
    """Create a .docx file with the full internship report."""

    # Content Types
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

    # Relationships
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

    # Word relationships
    word_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    # Styles
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="360" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
</w:styles>'''

    # Helper functions to build document XML
    def heading(text, level=1):
        style = f"Heading{level}"
        return f'''<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'''

    def title(text):
        return f'''<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'''

    def para(text, bold=False):
        bold_tag = "<w:b/>" if bold else ""
        return f'''<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr>{bold_tag}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'''

    def empty_para():
        return '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>'

    # Build document content
    body_content = []

    # Title Page
    body_content.append(empty_para())
    body_content.append(empty_para())
    body_content.append(title("Internship Report"))
    body_content.append(empty_para())
    body_content.append(f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t>Amul (Gujarat Cooperative Milk Marketing Federation Ltd.)</w:t></w:r></w:p>''')
    body_content.append(empty_para())
    body_content.append(f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Submitted by: [Student Name]</w:t></w:r></w:p>''')
    body_content.append(f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Roll No: [Roll Number]</w:t></w:r></w:p>''')
    body_content.append(f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Course: [Course Name]</w:t></w:r></w:p>''')
    body_content.append(f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Institution: [College/University Name]</w:t></w:r></w:p>''')
    body_content.append(f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Duration of Internship: [Start Date] to [End Date]</w:t></w:r></w:p>''')
    body_content.append(f'''<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Guide: [Faculty Guide Name]</w:t></w:r></w:p>''')
    body_content.append(empty_para())
    body_content.append(empty_para())

    # Page break
    body_content.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # Table of Contents
    body_content.append(heading("Table of Contents"))
    body_content.append(para("1. Introduction"))
    body_content.append(para("2. Company Profile"))
    body_content.append(para("3. Objectives of the Internship"))
    body_content.append(para("4. Methodology"))
    body_content.append(para("5. Findings"))
    body_content.append(para("6. Analysis"))
    body_content.append(para("7. Conclusion"))
    body_content.append(para("8. References"))
    body_content.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # 1. Introduction
    body_content.append(heading("1. Introduction"))
    body_content.append(para(
        "This report presents the findings and experiences gathered during the internship undertaken at Amul, "
        "officially known as the Gujarat Cooperative Milk Marketing Federation Ltd. (GCMMF). Amul is India's "
        "largest food brand and the world's largest pouched milk brand, with an annual turnover exceeding "
        "Rs. 72,000 crores. The organization operates on a unique three-tier cooperative structure that empowers "
        "millions of milk producers across Gujarat and beyond."
    ))
    body_content.append(para(
        "The internship was aimed at gaining practical exposure to the operations, marketing strategies, supply "
        "chain management, and organizational culture of one of India's most iconic dairy cooperatives. Over the "
        "course of the internship, the student was placed at [Department/Location] and had the opportunity to "
        "observe and participate in various business processes."
    ))
    body_content.append(para(
        "Amul's success story, often referred to as the 'White Revolution' or 'Operation Flood,' transformed India "
        "from a milk-deficient nation into the world's largest milk producer. The cooperative model pioneered by "
        "Dr. Verghese Kurien continues to serve as a benchmark for rural development and cooperative enterprise globally."
    ))
    body_content.append(empty_para())

    # 2. Company Profile
    body_content.append(heading("2. Company Profile"))
    body_content.append(heading("2.1 About Amul", level=2))
    body_content.append(para(
        "Amul (Anand Milk Union Limited) was established in 1946 in Anand, Gujarat, as a response to the exploitation "
        "of marginal milk producers by middlemen. The cooperative was formally registered on December 14, 1946, under "
        "the leadership of Tribhuvandas Patel, with active support from Sardar Vallabhbhai Patel. Dr. Verghese Kurien "
        "later joined as the general manager and transformed it into a world-class dairy organization."
    ))
    body_content.append(heading("2.2 Organizational Structure", level=2))
    body_content.append(para(
        "Amul operates through a three-tier cooperative structure: Village-level Dairy Cooperative Societies (DCS) "
        "at the base, District-level Milk Unions in the middle, and the Gujarat Cooperative Milk Marketing Federation "
        "(GCMMF) at the apex. This structure ensures that the benefits flow directly to the milk producers while "
        "maintaining operational efficiency at every level."
    ))
    body_content.append(heading("2.3 Product Portfolio", level=2))
    body_content.append(para(
        "Amul offers a diverse range of dairy products including milk, butter, cheese, ghee, ice cream, milk powder, "
        "chocolates, paneer, curd, buttermilk, flavored milk, and cream. The brand has also expanded into non-dairy "
        "segments such as potato-based snacks, edible oils, and bakery items. Amul's product line caters to various "
        "consumer segments across different price points."
    ))
    body_content.append(empty_para())

    # 3. Objectives
    body_content.append(heading("3. Objectives of the Internship"))
    body_content.append(para("The primary objectives of this internship were as follows:"))
    body_content.append(para("1. To understand the cooperative business model and its operational framework at Amul."))
    body_content.append(para("2. To study the supply chain management practices, from milk procurement to product distribution."))
    body_content.append(para("3. To analyze the marketing strategies and brand management techniques employed by Amul."))
    body_content.append(para("4. To examine the quality control and assurance processes in dairy production."))
    body_content.append(para("5. To gain practical knowledge about the day-to-day operations of a large-scale cooperative enterprise."))
    body_content.append(para("6. To identify key success factors and challenges faced by the organization in the competitive dairy industry."))
    body_content.append(empty_para())

    # 4. Methodology
    body_content.append(heading("4. Methodology"))
    body_content.append(heading("4.1 Research Approach", level=2))
    body_content.append(para(
        "The internship employed a combination of observational research, structured interviews, and secondary data "
        "analysis. A descriptive research design was adopted to capture the organizational processes, workflows, and "
        "strategic decision-making at Amul."
    ))
    body_content.append(heading("4.2 Data Collection Methods", level=2))
    body_content.append(para(
        "Primary Data: Direct observation of daily operations, informal and semi-structured interviews with managers "
        "and employees across departments, participation in team meetings and operational briefings, and hands-on "
        "involvement in assigned tasks."
    ))
    body_content.append(para(
        "Secondary Data: Company annual reports, published research papers on Amul's cooperative model, internal "
        "documents and presentations shared during orientation, industry reports from NDDB (National Dairy Development "
        "Board), and publicly available financial data."
    ))
    body_content.append(heading("4.3 Duration and Scope", level=2))
    body_content.append(para(
        "The internship spanned [X weeks/months] and was conducted at [specific plant/office/department]. "
        "The scope covered procurement operations, production processes, quality assurance, distribution logistics, "
        "and marketing activities. Weekly reports were maintained to document observations and learning outcomes."
    ))
    body_content.append(empty_para())

    # 5. Findings
    body_content.append(heading("5. Findings"))
    body_content.append(heading("5.1 Supply Chain Operations", level=2))
    body_content.append(para(
        "Amul's supply chain is one of the most efficient in the Indian FMCG sector. Milk is collected twice daily "
        "from over 3.6 million milk producer members across 18,700+ village-level dairy cooperative societies. "
        "The cold chain infrastructure ensures that milk reaches processing plants within hours of collection. "
        "Automated milk collection systems (AMCS) with electronic weighing, fat testing, and real-time data "
        "transmission have significantly improved transparency and efficiency at the village level."
    ))
    body_content.append(heading("5.2 Production and Quality Control", level=2))
    body_content.append(para(
        "The production facilities operate under stringent quality control protocols. Each batch undergoes multiple "
        "quality checks including fat content analysis, SNF (Solids-Not-Fat) measurement, adulteration detection, "
        "and microbiological testing. The organization follows ISO 22000 and HACCP standards. Quality is maintained "
        "at every stage from raw material intake to finished product dispatch, ensuring consistent product standards "
        "across all SKUs."
    ))
    body_content.append(heading("5.3 Marketing and Distribution", level=2))
    body_content.append(para(
        "Amul's distribution network spans across India with over 10,000 distributors and more than 10 lakh retail "
        "outlets. The 'Amul Girl' advertising campaign, running since 1966, is one of the longest-running ad campaigns "
        "globally. The marketing strategy combines competitive pricing (low-cost leadership), extensive distribution, "
        "strong brand recall, and topical advertising that keeps the brand culturally relevant. The company spends "
        "less than 1% of revenue on advertising, relying heavily on word-of-mouth and product quality."
    ))
    body_content.append(heading("5.4 Human Resource Practices", level=2))
    body_content.append(para(
        "Amul maintains a lean organizational structure despite its massive scale. Employee engagement is driven by "
        "a sense of purpose — working for a cooperative that benefits millions of rural farmers. Training programs "
        "are conducted regularly for both cooperative society members and professional staff. The organization promotes "
        "a culture of innovation, allowing employees to experiment with new product formulations and process improvements."
    ))
    body_content.append(heading("5.5 Technology and Innovation", level=2))
    body_content.append(para(
        "Amul has embraced technology at multiple levels: ERP systems for enterprise management, automated milk "
        "collection units at village societies, IoT-enabled cold chain monitoring, and data analytics for demand "
        "forecasting and route optimization. The organization also invests in R&D for developing new products "
        "and improving shelf life of existing products."
    ))
    body_content.append(empty_para())

    # 6. Analysis
    body_content.append(heading("6. Analysis"))
    body_content.append(heading("6.1 SWOT Analysis", level=2))
    body_content.append(para("Strengths:", bold=True))
    body_content.append(para(
        "- Strong cooperative structure ensuring farmer loyalty and consistent raw material supply. "
        "- Extensive product portfolio catering to diverse consumer needs. "
        "- Iconic brand with high recall and trust among Indian consumers. "
        "- Efficient cold chain and distribution infrastructure. "
        "- Low-cost business model enabling competitive pricing."
    ))
    body_content.append(para("Weaknesses:", bold=True))
    body_content.append(para(
        "- Dependence on Gujarat for majority of milk procurement. "
        "- Limited presence in premium/organic dairy segment compared to private competitors. "
        "- Cooperative decision-making can sometimes be slow. "
        "- Limited international brand presence relative to its production capacity."
    ))
    body_content.append(para("Opportunities:", bold=True))
    body_content.append(para(
        "- Growing health-conscious consumer base demanding value-added dairy products. "
        "- Expansion into international markets with Indian diaspora as initial target. "
        "- E-commerce and direct-to-consumer channels for premium products. "
        "- Plant-based dairy alternatives and functional foods segment."
    ))
    body_content.append(para("Threats:", bold=True))
    body_content.append(para(
        "- Increasing competition from private dairy companies (Mother Dairy, Nestle, Britannia). "
        "- Climate change affecting milk production and fodder availability. "
        "- Regulatory changes in food safety and pricing. "
        "- Shifting consumer preferences towards non-dairy alternatives."
    ))
    body_content.append(heading("6.2 Competitive Advantage Analysis", level=2))
    body_content.append(para(
        "Amul's primary competitive advantage lies in its cooperative ownership model, which eliminates intermediaries "
        "and ensures that approximately 80% of the consumer price reaches the milk producers. This contrasts sharply "
        "with private dairy companies where farmer realization is typically 50-60%. This model creates a self-reinforcing "
        "cycle: better prices attract more farmers, more milk enables economies of scale, lower costs allow competitive "
        "pricing, higher volumes lead to better farmer prices."
    ))
    body_content.append(heading("6.3 Key Performance Indicators", level=2))
    body_content.append(para(
        "During the internship period, the following KPIs were observed: Daily milk procurement averaging 26+ million "
        "liters per day, plant capacity utilization exceeding 85%, product rejection rate below 0.5%, distribution "
        "reach covering 50+ million households, and consistent year-on-year revenue growth of 15-20%. These metrics "
        "demonstrate the operational excellence and market strength of the cooperative."
    ))
    body_content.append(empty_para())

    # 7. Conclusion
    body_content.append(heading("7. Conclusion"))
    body_content.append(para(
        "The internship at Amul provided invaluable insights into the functioning of India's largest dairy cooperative. "
        "The experience highlighted how a well-structured cooperative model can achieve commercial success while "
        "simultaneously uplifting millions of rural livelihoods. Amul's ability to balance social objectives with "
        "business efficiency serves as a powerful case study in sustainable and inclusive business practices."
    ))
    body_content.append(para(
        "Key takeaways from the internship include: (1) the importance of backward integration and farmer-centric "
        "supply chain design, (2) how consistent branding and low-cost marketing can build enduring consumer trust, "
        "(3) the role of technology in scaling cooperative operations without compromising quality, and (4) the "
        "significance of institutional design in ensuring democratic governance and equitable benefit-sharing."
    ))
    body_content.append(para(
        "The internship also revealed challenges that Amul faces going forward, including increasing private sector "
        "competition, changing consumer preferences, and the need for continued technological investment. However, "
        "the organization's strong cooperative foundations, brand equity, and operational capabilities position it "
        "well to navigate these challenges and continue its growth trajectory."
    ))
    body_content.append(para(
        "In conclusion, the internship was a highly enriching experience that bridged theoretical knowledge with "
        "practical industry exposure. The lessons learned about cooperative management, supply chain excellence, "
        "and brand building will be valuable for future professional endeavors."
    ))
    body_content.append(empty_para())

    # 8. References
    body_content.append(heading("8. References"))
    body_content.append(para("1. Amul Official Website - www.amul.com"))
    body_content.append(para("2. Gujarat Cooperative Milk Marketing Federation (GCMMF) Annual Report 2023-24."))
    body_content.append(para("3. Kurien, V. (2005). I Too Had a Dream. Roli Books."))
    body_content.append(para("4. National Dairy Development Board (NDDB) - www.nddb.coop"))
    body_content.append(para("5. Scholten, B. (2010). India's White Revolution. Taurus Academic Studies."))
    body_content.append(para("6. Chandra, P. (2014). Building the Amul Business: Cooperative Enterprise. IIM Ahmedabad Case Study."))
    body_content.append(para("7. Ministry of Fisheries, Animal Husbandry and Dairying, Government of India - Annual Report."))

    # Assemble document.xml
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {''.join(body_content)}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    # Create the .docx file
    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('word/_rels/document.xml.rels', word_rels)
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/styles.xml', styles)

    print(f"Report created successfully: {filename}")
    print(f"File size: {os.path.getsize(filename)} bytes")


if __name__ == "__main__":
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Amul_Internship_Report.docx")
    create_docx(output_path)
