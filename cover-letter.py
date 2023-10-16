from docx import Document


def generate_cover_letter(company_name):
    cover_letter = f"""
    Dear Hiring Manager,

    I am writing to apply for the open position at {company_name}. I have a strong background in [mention your relevant skills or experiences], and I am excited about the opportunity to contribute my expertise to your team.

    [Customize the content of the cover letter as needed.]

    Thank you for considering my application. I look forward to the possibility of discussing how my skills and experience can benefit {company_name}.

    Sincerely,
    [Your Name]
    """
    return cover_letter


def create_word_doc(text, file_name):
    doc = Document()
    doc.add_heading('Michael Gulson')
    doc.add_paragraph()
    doc.add_paragraph(text)
    doc.save(f"{file_name}.docx")



default_company = "Company"  # Default company name
company_name = input(f"Enter the company name (press enter to use the default {default_company}): ")
company_name = company_name.strip() if company_name else default_company

cover_letter_text = generate_cover_letter(company_name)
file_name = f"Cover_Letter(10-16-23){company_name.replace(' ', '_')}"
create_word_doc(cover_letter_text, file_name)
print(f"{file_name}.docx created successfully.")
