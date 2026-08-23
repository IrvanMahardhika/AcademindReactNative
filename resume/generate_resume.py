#!/usr/bin/env python3
"""Generate ATS-friendly Senior Full-Stack resume PDF."""

from fpdf import FPDF


class ResumePDF(FPDF):
    def header(self):
        pass

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 8, f"Page {self.page_no()}", align="C")


def section_title(pdf: ResumePDF, title: str) -> None:
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(0, 7, title.upper(), new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(60, 60, 60)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(3)


def bullet(pdf: ResumePDF, text: str) -> None:
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(40, 40, 40)
    pdf.multi_cell(0, 5, f"  - {text}")


def tech_line(pdf: ResumePDF, text: str) -> None:
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "I", 9.5)
    pdf.set_text_color(70, 70, 70)
    pdf.multi_cell(0, 5, f"  {text}")


def job_header(pdf: ResumePDF, title: str, company: str, dates: str, location: str) -> None:
    pdf.ln(1)
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(0, 5, title, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 5, f"{company} | {location} | {dates}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1)


def build_resume(output_path: str) -> None:
    pdf = ResumePDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.set_margins(18, 16, 18)
    pdf.add_page()

    # Name & headline
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(15, 15, 15)
    pdf.cell(0, 9, "Irvan Mahardhika Setyawan", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(45, 45, 45)
    pdf.cell(0, 6, "Senior Full-Stack Engineer", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    pdf.set_font("Helvetica", "", 9.5)
    pdf.set_text_color(50, 50, 50)
    contact_line_1 = (
        "Bogor, Indonesia  |  +62 812-1039-4457  |  irvan.mahardhika@gmail.com"
    )
    contact_line_2 = (
        "https://www.linkedin.com/in/irvan-mahardhika/  |  https://github.com/IrvanMahardhika"
    )
    pdf.multi_cell(0, 5, contact_line_1)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(0, 5, contact_line_2)
    pdf.ln(2)

    # Summary
    section_title(pdf, "Professional Summary")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(40, 40, 40)
    summary = (
        "Senior Full-Stack Engineer with 6+ years of experience building scalable web and mobile "
        "applications across fintech, logistics, and e-commerce. Proven track record leading "
        "engineering teams, designing production systems with React, Vue, NestJS, and AWS, and "
        "delivering high-impact features including Shopify integrations, AI-powered OCR, and "
        "real-time logistics platforms. Strong in API design, cloud infrastructure, code quality, "
        "and cross-functional collaboration."
    )
    pdf.multi_cell(0, 5, summary)

    # Skills
    section_title(pdf, "Technical Skills")
    skills = [
        ("Languages", "JavaScript, TypeScript, Python, SQL"),
        ("Frontend", "React.js, Vue.js, Next.js, React Native, Expo, Angular, HTML, CSS"),
        ("Backend", "NestJS, Node.js, Express.js, Django REST Framework, GraphQL, REST APIs"),
        ("Databases", "PostgreSQL, MySQL, Redis, DynamoDB, MongoDB (NoSQL)"),
        (
            "Cloud & DevOps",
            "AWS (S3, EC2, ECS, Lambda, API Gateway, RDS, CloudFront), Docker, Kubernetes, GCP",
        ),
        (
            "Tools & Practices",
            "Git, CI/CD, BullMQ, TypeORM, Shopify, OpenCV, PaddleOCR, Agile, Code Review",
        ),
    ]
    for label, value in skills:
        pdf.set_x(pdf.l_margin)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(40, 40, 40)
        pdf.multi_cell(0, 5, f"{label}: {value}")
        pdf.ln(0.5)

    # Experience
    section_title(pdf, "Professional Experience")

    job_header(
        pdf,
        "Senior Full-Stack Engineer",
        "Wayfindr",
        "Feb 2026 - Present",
        "Singapore (Full-time)",
    )
    for item in [
        "Built and maintained a multi-app Vue.js client application with Vite for production use across client workflows.",
        "Developed NestJS backend services with TypeORM, multi-database PostgreSQL, Redis, BullMQ, and real-time features.",
        "Built and operated Shopify connectivity (OAuth, GraphQL order sync, fulfillment flows) and drove Shopify App Store listing readiness.",
        "Led AI-assisted technical design for a greenfield Freight Shipment Management module, translating PRD into architecture, DB schema, API specs, and dynamic UI/UX wireframes.",
    ]:
        bullet(pdf, item)
    tech_line(pdf, "Tech: Vue.js, NestJS, Docker, PostgreSQL, Redis, Shopify, TypeORM, BullMQ")

    job_header(
        pdf,
        "Software Engineer (Freelance)",
        "InstaRinn",
        "Dec 2025 - Feb 2026",
        "India",
    )
    for item in [
        "Built an SEO-friendly Next.js web application optimized for search engine discoverability, fast load times, and seamless indexing.",
        "Developed serverless-style backends on Supabase for scalable data and authentication workflows.",
        "Delivered a React.js PWA with service workers for offline functionality and a native-like user experience.",
    ]:
        bullet(pdf, item)
    tech_line(pdf, "Tech: Next.js, React.js, Supabase, PWA, SEO")

    job_header(
        pdf,
        "Software Engineer (Freelance)",
        "Cloudetica Solutions",
        "Sep 2025 - Jan 2026",
        "Indonesia",
    )
    for item in [
        "Developed a PWA using React Native and Expo, integrated with a NestJS backend supporting 1K+ users and 50K+ requests per day.",
        "Built interactive web applications using React.js for client-facing product features.",
        "Deployed and managed cloud infrastructure on AWS including S3, EC2, ECR, ECS, ELB, DynamoDB, RDS, and CloudFront.",
        "Conducted thorough code reviews to maintain code quality and team standards.",
    ]:
        bullet(pdf, item)
    tech_line(pdf, "Tech: React Native, Expo, React.js, NestJS, PostgreSQL, NoSQL, GraphQL, AWS")

    job_header(
        pdf,
        "Senior Software Engineer",
        "TVS Digital",
        "Oct 2023 - Oct 2025",
        "Singapore (Contract)",
    )
    for item in [
        "Led a team of 3 software engineers in full-stack development using React.js, React Native, AWS Lambda (serverless), and Express.js.",
        "Built an AI-powered OCR system to extract text from IDs using OpenCV and PaddleOCR, wrapped in a Django REST Framework service.",
        "Deployed and managed AWS infrastructure including S3, Lambda, API Gateway, RDS, and CloudFront.",
        "Spearheaded production support operations and resolved 50+ L3 support tickets per month.",
        "Led delivery of a mission-critical software project, resulting in a long-term client partnership and increased revenue.",
        "Received 2024 TVSD Spot Award for exemplary performance.",
    ]:
        bullet(pdf, item)
    tech_line(pdf, "Tech: React.js, React Native, Node.js, Express.js, Django REST Framework, MySQL, AWS Lambda")

    job_header(
        pdf,
        "Senior Software Engineer",
        "Danamas",
        "Aug 2022 - Aug 2023",
        "Indonesia (Contract)",
    )
    for item in [
        "Directed a team of software engineers in frontend development using React Native and Angular.",
        "Published and maintained applications on Apple App Store and Google Play Store.",
        "Conducted code reviews and upheld engineering best practices across the mobile codebase.",
    ]:
        bullet(pdf, item)
    tech_line(pdf, "Tech: React Native, Angular, Xcode, Android Studio, Visual Studio App Center")

    job_header(
        pdf,
        "Software Engineer",
        "Danacita",
        "Apr 2021 - Aug 2022",
        "Indonesia (Full-time)",
    )
    for item in [
        "Built web applications with React.js and Wagtail CMS.",
        "Built mobile applications with React Native and Expo.",
        "Developed RESTful APIs with Django REST Framework.",
        "Executed deployments across development, staging, and production environments.",
        "Investigated production errors, identified root causes, and implemented fixes.",
        "Recognized as a key employee and awarded stock options.",
    ]:
        bullet(pdf, item)
    tech_line(pdf, "Tech: React.js, React Native, Expo, Wagtail, Django REST, Docker, Kubernetes, GCP")

    job_header(
        pdf,
        "Software Engineer",
        "Purwadhika Digital Technology School",
        "Jan 2020 - Mar 2021",
        "Indonesia (Contract)",
    )
    for item in [
        "Built mobile applications with React Native and published to Apple App Store and Google Play Store.",
        "Built web applications with React.js and RESTful APIs with Node.js and Express.js.",
        "Investigated production errors and delivered timely resolutions.",
    ]:
        bullet(pdf, item)
    tech_line(pdf, "Tech: React.js, React Native, Node.js, Express.js, MySQL")

    # Prior Experience
    section_title(pdf, "Prior Experience")
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(40, 40, 40)
    pdf.multi_cell(
        0,
        5,
        "Retail Employee | Superindo, Ace Hardware, Central Retail | 2005 - 2019",
    )

    # Education
    section_title(pdf, "Education")
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(0, 5, "Bachelor's Degree, Business Administration and Management", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 5, "IPB University | Aug 2001 - Sep 2005", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(0, 5, "Full-Stack Web and Mobile Development", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 5, "Purwadhika Digital Technology School | Jul 2019 - Dec 2019", new_x="LMARGIN", new_y="NEXT")

    pdf.output(output_path)


if __name__ == "__main__":
    paths = [
        "/workspace/resume/Irvan_Mahardhika_Setyawan_Senior_FullStack_Resume.pdf",
        "/opt/cursor/artifacts/Irvan_Mahardhika_Setyawan_Senior_FullStack_Resume.pdf",
    ]
    for path in paths:
        build_resume(path)
        print(f"Created: {path}")
