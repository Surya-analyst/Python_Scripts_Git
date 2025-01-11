from Resume_Generator_Custom_Definitions import *

#filename = "Surya_Resume_Updated_On_" + str(date.today()) + ".docx"
Pdf_Name = "Surya_Resume_Updated_On_" + str(date.today()) + ".pdf"

#define Perosnal Details
Name = "XXXXX XXXXX" #XXXXX XXXXX
phone_symbol = "\U0001F4DE"
Email_symbol = "\U0001F4E7"
url_symbol = "\U0001F517"
Contact_Number = "(+XX) XXXXX XXXXX" #(+XX) XXXXX XXXXX
Email_ID = "XXXXXXXXXX@gmail.com" #XXXXXXXXXX
Email_link = "mailto:" + Email_ID
LinkedIn = "www.linkedin.com" #"www.linkedin.com"
LinkedIn_URL = "https://" + LinkedIn
Contact_Number_link = Contact_Number
Contact_Number_link = Contact_Number_link.replace("(","")
Contact_Number_link = Contact_Number_link.replace(") ","")
Contact_Number_link = Contact_Number_link.replace(" ","")
Contact_Number_link = "tel: //" + Contact_Number_link

#declare document color
doc_color = '#FEFDF4'

def create_resume():
    # Create a new Word Document
    doc = Document()
   
    # Set document properties
    core_properties = doc.core_properties
    core_properties.title = 'Professional Resume of ' + Name
    core_properties.author = Name
    core_properties.subject = 'Resume'
    
    #Set document layout, background colours, etc
    document_colour(doc, doc_color)
    
    # Define a custom style for name heading
    heading_style = doc.styles.add_style('CustomHeading', WD_STYLE_TYPE.PARAGRAPH)
    heading_style.font.name = 'Calibri'
    heading_style.font.size = Pt(16)
    heading_style.font.color.rgb = RGBColor(0, 0, 0) 
    heading_style.font.bold = True
    heading_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    
    #set header & Adjust page margins (top and bottom)
    
    section = doc.sections[0] #Choosing the top most section of the page
    section.top_margin = Pt(60)  #Reduce the top margin for header
    section.bottom_margin = Pt(0.7)  #Reduce the bottom margin for footer
     
    # Selecting the header
    header = section.header
        
    # Add content to header
    header_para = header.paragraphs[0]
    header_para.text = ">>> This resume was generated entirely in Python. For full sourcecode "
    header_custmizaion(header_para)
    d = add_hyperlink(header_para,'Click Here' ,'https://github.com/Surya-analyst/Python_Scripts_Git/blob/main/Resume_Generator_v2_for_git.py')
 
    # Add header with name and contact details
    
    doc.add_heading(Name, level=1).style = 'CustomHeading'
    Skills = doc.add_paragraph('Data Analyst | Excel | VBA | SQL | Python | Power BI | Tableau')
    Level1_customization(Skills)
    
    #Add contact details
    contact_paragraph = doc.add_paragraph()
    psym_run = contact_paragraph.add_run(phone_symbol)
    symbol_customization(psym_run)
    add_hyperlink(contact_paragraph," " + Contact_Number + " | ", Contact_Number_link)
    esymbol_run = contact_paragraph.add_run(Email_symbol)
    symbol_customization(esymbol_run)
    add_hyperlink(contact_paragraph," " + Email_ID + " | ", Email_link)
    usymbol_run = contact_paragraph.add_run(url_symbol)
    symbol_customization(usymbol_run)
    add_hyperlink(contact_paragraph," " + LinkedIn , LinkedIn_URL)
    contact_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    paragraph_format = contact_paragraph.paragraph_format
    paragraph_format.space_after = 0  # Minimize space after the paragraph after the line
    paragraph_format.space_before = 0 # Minimize space before the paragraph before the line
    
    #to insert horizontal line
    doc.add_paragraph()
    add_horizontal_line(doc)
     
    # Add Professional Summary
    doc.add_heading('Professional Summary', level=2)
    add_para(doc,'','Results-driven Data Analyst with 6 years of experience in leveraging data to drive actionable insights and strategic decisions. Proficient in Data Fetching, Data cleansing, data preparation and data visualization using tools such as SQL, Excel, VBA, Python, Tableau and PowerBI. Adept at collaborating with cross-functional teams to identify key business challenges and implement data-driven solutions that optimize performance.')
    
    #to insert horizontal line
    doc.add_paragraph()
    add_horizontal_line(doc)
  
    # Add Core Skills
    a = doc.add_heading('Core Skills', level=2)
    add_bullet_point(doc, '', 'Data Analysis & Reporting (SQL, Excel, Tableau, PowerBI)')
    add_bullet_point(doc, '', 'Data Automation & Process Optimization (VBA, Python, Power Query)')
    add_bullet_point(doc, '', 'Project Coordination & Documentation')
    add_bullet_point(doc, '', 'Stakeholder Communication & Collaboration')
    
    #to insert horizontal line
    doc.add_paragraph()
    add_horizontal_line(doc)
   
    # Add Skill rating section with bar graph
    doc.add_heading('Skill Rating', level=2)
    skills = ["Excel", "MySQL", "VBA", "Python", "Power BI", "Tableau"]
    ratings = [5, 4, 3, 3, 3, 3]
    max_rating = 5  # Maximum possible rating

    # Colors for the bars
    bar_colors = ['#009879'] * len(skills)  # Main bar color
    bg_colors = ['#b7e4c7'] * len(skills)  # Background bar color

    # Plot settings
    plt.figure(figsize=(4, 2))
    bar_height = 0.7  # Adjust height to reduce spacing (default is typically 0.8)
    for i, (skill, rating) in enumerate(zip(skills, ratings)):
        # Draw background bar
        plt.barh(y=i, width=max_rating, color=bg_colors[i], height=bar_height, edgecolor='none')
        # Draw rating bar
        plt.barh(y=i, width=rating, color=bar_colors[i], height=bar_height, edgecolor='none')
        # Add text for the rating and skill
        plt.text(rating - 0.9, i, f"{rating} / {max_rating}", va='center', ha='center', color='white', fontsize=11, weight='bold', fontname = 'Calibri')
        plt.text(-0.99, i, skill, va='center', ha='left', color='black', fontsize=11, fontname = 'Calibri')

    # Remove axes and labels for clean look
    plt.gca().invert_yaxis()  # Invert y-axis for descending order
    plt.axis('off')  # Hide the axes

    # Set background color
    plt.gcf().patch.set_facecolor(doc_color)  # match document color

    # Show or Save the chart as an image
    plt.tight_layout()
    #plt.show()
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
        chart_path = tmp_file.name
        plt.savefig(chart_path)
    plt.close()
    
    # Add the chart to the Word document
    doc.add_picture(chart_path, width=doc.sections[0].page_width * 0.4)  # Scale image to fit page
    #to insert horizontal line
    add_horizontal_line(doc)
    
    # Add Professional Experience
    doc.add_heading('Professional Experience', level=2)
    doc.add_paragraph('Assistant Manager - MIS', style='Heading 3')
    add_company_details(doc,'Standard Chartered Bank - GBS, ', 'Chennai','\t' * 4 + ' ' * 9 + '07/2023 – Present', 'A global bank with comprehensive financial services.')
    add_bullet_point(doc, '','Automate daily report consolidation and data extraction using Power Query, improving data accuracy and efficiency.')
    add_bullet_point(doc, '', 'Develop solutions with Excel User Forms, VBA, and Power Query, enhancing process speed and reliability.')
    add_bullet_point(doc, '', 'Create dynamic tools for seamless data management between MS Access and Excel.')
    add_bullet_point(doc, '', 'Implement automated reporting for trend analysis and productivity, enabling data-driven decision-making.')
  
    doc.add_page_break()
    #doc.add_paragraph('')
    a = doc.add_paragraph('Data Analyst', style='Heading 3')
    add_company_details(doc,'Optum Health Care Business Services & Technology, ','Chennai','\t' * 2 + ' ' * 8 +'03/2021 – 07/2023', 'A healthcare technology company providing services to improve healthcare delivery.')
    #add_bullet_point(doc, '', '')
    add_bullet_point(doc, '', 'Extract, clean, and analyze large datasets, ensuring quality data for business insights.')
    add_bullet_point(doc, '', 'Led project planning, task scheduling, and documentation for healthcare solutions.')
    add_bullet_point(doc, '', 'Coordinate with cross-functional teams to deliver client projects successfully.')
    add_bullet_point(doc, '', 'Develop SQL queries and automate reporting to support management decisions.')
    
    doc.add_paragraph('Technical Process Specialist', style='Heading 3')
    add_company_details(doc,'Infosys, ','Bangalore','\t' * 7 + ' ' * 8 + '09/2020 – 03/2021', 'A multinational corporation that provides business consulting, information technology, and outsourcing services.')
    #add_bullet_point(doc, '', '')
    add_bullet_point(doc, '', 'Extract and analyze data from CMS platforms for custom reporting.')
    add_bullet_point(doc, '', 'Build Excel dashboards and manage SQL databases for data operations.')
    add_bullet_point(doc, '', 'Deliver stakeholder reports, ensuring clear communication of insights.')

    doc.add_paragraph('Project Coordinator', style='Heading 3')
    add_company_details(doc,'Origin Learning Solutions Pvt Ltd, ','Chennai','\t' * 4 + ' ' * 8 + '10/2018 – 01/2020', 'An organization focused on learning solutions and educational content.')
    #add_bullet_point(doc, '', '')
    add_bullet_point(doc, '', 'Manage project schedules, resources, and milestones for smooth execution.')
    add_bullet_point(doc, '', 'Collaborate with stakeholders to define objectives and monitor progress.')
    add_bullet_point(doc, '', 'Develop risk assessments and contingency plans to mitigate project issues.')
    
    doc.add_paragraph('Project Associate', style='Heading 3')
    add_company_details(doc,'Emerson Automation Solutions Pvt Ltd, ','Chennai','\t' * 3 + ' ' * 8 + '07/2017 – 07/2018', 'A company specializing in automation solutions for a variety of industries.')
    #add_bullet_point(doc, '', '')
    add_bullet_point(doc, '', 'Oversee PMO tasks, delivery schedules, and technical document submissions.')
    add_bullet_point(doc, '', 'Maintain client relationships by addressing feedback and meeting deadlines.')
    
    #to insert horizontal line
    doc.add_paragraph()
    add_horizontal_line(doc)
   
    # Add Education section
    doc.add_heading('Education', level=2)
    c = doc.add_paragraph('Bachelor of Engineering', style='Heading 3')
    add_para(doc,'','R.M.D Engineering College (2017)')
    # Add an indent to the bullet
    p_format = c.paragraph_format
    p_format.left_indent = Cm(1.0)  # Indent bullet by 1 cm (adjust as needed)

    #to insert horizontal line
    doc.add_paragraph()
    add_horizontal_line(doc)
    
    doc.add_heading('Additional Training or Certifications', level=2)
    c = doc.add_paragraph('MySQL for Data Analytics and Buisness Intelligence (2020), ', style='List Bullet')
    d = add_hyperlink(c,'Certificate Link' ,'https://www.udemy.com/certificate/UC-3c8b12ac-16ab-441c-a5b8-a71f287c31c1/')
    c = Custom_Style(c)

    c = doc.add_paragraph('Master Python by Coding 100 Practical Problems (2024), ', style='List Bullet')
    d = add_hyperlink(c,'Certificate Link' ,'https://www.udemy.com/certificate/UC-f4484a94-e2bb-42b9-80f8-efff2f20ae24/')
    c = Custom_Style(c)
    
    #to insert horizontal line
    doc.add_paragraph()
    add_horizontal_line(doc)
    

    # Add Languages (Optional)  
    doc.add_heading('Languages', level=2)
    c = doc.add_paragraph('Tamil (Proficient)', style='List Bullet')
    d = doc.add_paragraph('English (Proficient)', style='List Bullet')
    c = Custom_Style(c)
    d = Custom_Style(d)

    # Save the document
    # doc.save(filename)
    
    # Save the document as temp file to generate pdf
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
        filename = tmp_file.name
        doc.save(filename)
    plt.close()
    
    #convert the docx temp file to pdf
    convert(filename, Pdf_Name)
    
# Call the function to create the resume
create_resume()