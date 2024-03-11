while True:
    initializer = input('This application prepares born-digital DIPs for deposit. Choose an option:\n\n1.   PrependPDF for DRS only\n2.   PrependPDF for VRR only\n3.   Package converted files for VRR (NO PDF)\n')

    if initializer == '1':
        folder_path = input('Enter the file path for the folder containing PDFs, your file list CSV, and the PUI finding aid CSV.\n')
        VRR_check = False
        VRR_PDF_check = False
    elif initializer == '2':
        folder_path = input('Enter the file path for the folder containing PDFs, your file list CSV, and the PUI finding aid CSV.\n')
        VRR_check = False
        VRR_PDF_check = True
    elif initializer == '3':
        folder_path = input('Enter the file path for the folder containing converted files, your file list CSV, and the PUI finding aid CSV.\n')
        VRR_check = True
        VRR_PDF_check = False

    import csv
    import pdfkit
    import os
    from os import listdir
    from os.path import isfile, join
    import shutil
    import fitz
    import PyPDF2
    import pandas as pd
    from PyPDF2 import PdfReader
    from PyPDF2 import PdfMerger
    from airium import Airium
    from reportlab.pdfgen.canvas import Canvas
    from pdfrw import PdfReader
    from pdfrw.toreportlab import makerl
    from reportlab.lib.pagesizes import letter
    from pdfrw.buildxobj import pagexobj
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import calendar
    import img2pdf


    options = {
        'page-size': 'Letter',
        'margin-top': '20',
        'margin-right': '20',
        'margin-bottom': '20',
        'margin-left': '20',
        'encoding': "UTF-8",
        'custom-header': [
            ('Accept-Encoding', 'gzip')
        ],
        'no-outline': None,
        # 'characterSpacing': '30',
        'footer-right': '[page]',
        'page_size': 'A4',
        'print_media_type': True,
        'dpi': '600'
    }

    def create_item_cover():
        a = Airium()
        a('<!DOCTYPE html>')
        with a.html(lang="pl"):
            with a.head():
                a.meta(charset="utf-8-sig")
                a.meta(name="pdfkit-page-size", content="Letter")
                a.meta(name="pdfkit-margin-top", content="10")
                a.meta(name="pdfkit-margin-right", content="20")
                a.meta(name="pdfkit-margin-left", content="20")
                a.meta(name="pdfkit-margin-bottom", content="10")
                a.title(_t="MC-###")
            with a.div(height='800'):
                with a.body():
                    with a.div(id = 'page-container'):
                        with a.div(width="100%", id = 'content-wrap'):
                            with a.table(height="125", width='100%', id='table_header'):
                                with a.tr(klass='header_row'):
                                    with a.td(width='325', style="vertical-align: top;"):
                                        a.img(width='270', style="vertical-align: top;", alt='logo_2022', src='https://i.imgur.com/prr5zvO.png')
                                    a.td(width='50', _t='')
                                    a.td(width='500', style="vertical-align: top; text-align: right", _t='Schlesinger Library<br/>Harvard Radcliffe Institute<br/>3 James St<br/>Cambridge, MA 02138<br/>radcliffe.harvard.edu/schlesinger-library')
                            a.hr()
                            with a.p():
                                a("This cover page describes the file \"" + Name + ".\" The file's contents begin on the following page. All times are given in UTC and in the 24-hour format.")
                            with a.h1(id="id00000001", kclass='file_header'):
                                a('File name:')
                                a(Name)
                            with a.table(id='File_info'):
                                with a.tr(klass='no_header'):
                                    a.td(width='215', style="font-weight:bold; vertical-align:top; font-size: 1.2em", _t="Original file category:")
                                    a.td(style="vertical-align:top; font-size: 1.2em", _t=Category)
                                with a.tr():
                                    a.td(style="font-weight:bold; vertical-align:top; font-size: 1.2em", _t="File size on disk:")
                                    a.td(style="vertical-align:top; font-size: 1.2em", _t=Size)
                                with a.tr():
                                    a.td(style="font-weight:bold; vertical-align:top; font-size: 1.2em", _t="Page count:")
                                    a.td(style="vertical-align:top; font-size: 1.2em", _t=Page_count)
                                with a.tr():
                                    a.td(style="font-weight:bold; vertical-align:top; font-size: 1.2em", _t="Date created:")
                                    a.td(style="vertical-align:top; font-size: 1.2em", _t=Created_date)
                                with a.tr():
                                    a.td(style="font-weight:bold; vertical-align:top; font-size: 1.2em", _t="Last modified:")
                                    a.td(style="vertical-align:top; font-size: 1.2em", _t=Modified_date)
                                with a.tr():
                                    a.td(style="font-weight:bold; vertical-align:top; font-size: 1.2em", _t="Author:")
                                    a.td(style="vertical-align:top; font-size: 1.2em", _t=Author)
                                with a.tr():
                                    a.td(style="font-weight:bold; vertical-align:top; font-size: 1.2em", _t="Last saved by:")
                                    a.td(style="vertical-align:top; font-size: 1.2em", _t=Last_saved_by)
        html = str(a)
        stylesheet_path = "C:\Temp\PrependPDF-css\\"
        css = stylesheet_path + 'stylesheet.css'
        pdfkit.from_string(html, folder_path+'\\'+Item_number+'_cover'+'.pdf', css=css)

    def create_master_cover():
        a2 = Airium()
        a2('<!DOCTYPE html>')
        with a2.html(lang="pl"):
            with a2.head():
                a2.meta(charset="utf-8-sig")
                a2.meta(name="pdfkit-page-size", content="Letter")
                a2.meta(name="pdfkit-margin-top", content="10")
                a2.meta(name="pdfkit-margin-right", content="20")
                a2.meta(name="pdfkit-margin-left", content="20")
                a2.meta(name="pdfkit-margin-bottom", content="10")
                a2.title(_t="MC-###")
            with a2.div(height='800'):
                with a2.body():
                    with a2.div(id = 'page-container'):
                        with a2.div(width="800", id = 'content-wrap'):
                            with a2.table(height="125", width='100%', id='table_header'):
                                with a2.tr(klass='header_row'):
                                    with a2.td(width='100', style="vertical-align: top;"):
                                        a2.img(width='270', style="vertical-align: top;", alt='logo_2022', src='https://i.imgur.com/prr5zvO.png')
                                    a2.td(width='50', _t='')
                                    a2.td(width='500', style="vertical-align: top; text-align: right", _t='Schlesinger Library<br/>Harvard Radcliffe Institute<br/>3 James St<br/>Cambridge, MA 02138<br/>radcliffe.harvard.edu/schlesinger-library')
                            a2.hr()
                            a2.br()
                            with a2.table(id='Collection_info'):
                                with a2.tr(klass='no_header'):
                                    a2.td(width='200', style="font-weight:bold; vertical-align:top", _t='Collection title:')
                                    a2.td(style="vertical-align:top", _t=Collection_name)
                                with a2.tr():
                                    a2.td(style="font-weight:bold; vertical-align:top", _t='Collection identifier:')
                                    a2.td(style="vertical-align:top", _t=Collection_ID)
                                with a2.tr():
                                    a2.td(style="font-weight:bold; vertical-align:top", _t='Item title:')
                                    a2.td(style="vertical-align:top", _t=Folder_name)
                                with a2.tr():
                                    a2.td(style="font-weight:bold; vertical-align:top", _t='Item identifier:')
                                    a2.td(style="vertical-align:top", _t=E_number)
                                with a2.tr():
                                    a2.td(style="font-weight:bold; vertical-align:top", _t='Finding aid permalink:')
                                    a2.td(style="vertical-align:top", _t=Permalink)
                            a2.br()
                            with a2.p():
                                a2('This download contains '+item_count+' files from item '+E_number+'. The original data has been converted to PDF for distribution and access.')
                            a2.br()
                            with a2.p():
                                a2('Please use the bookmark navigation pane in your PDF viewer to navigate this document. Each bookmark will direct you to a file in this folder. Schlesinger Library has added a cover page to each file containing metadata extracted from the original file, such as the file name and size.')
                            a2.br()
                            with a2.p():
                                a2.t(style="font-weight:bold", _t='Reading file anomalies:')
                            with a2.p():
                                a2('In files where the Created date is later than the Modified date, this apparent discrepancy is unlikely to be an actual error. It is more likely that the file was copied after its last modification, resulting in a new Created date but leaving the correct Modified date.')
                            a2.br()
                            with a2.p():
                                a2('Due to potential problems during conversion (e.g. file corruption or character encoding), it is possible that some pages display strange symbols instead of text.')
                            a2.br()
                            with a2.p():
                                a2.t(style="font-weight:bold", _t='Preferred citation:')
                            with a2.p():
                                a2(Collection_name+'.')
                                a2(Collection_ID+',')
                                a2('item ')
                                a2(E_number+'.')
                                a2('Schlesinger Library, Radcliffe Institute, Harvard University, Cambridge, Mass.')
                            a2.br()
                            with a2.p():
                                a2.t(style="font-weight:bold", _t='Rights:')
                            with a2.p():
                                a2('Conditions governing the use of this document may apply. For more information see the section entitled "Conditions Governing Use" in the collection finding aid linked above.')
                            a2.br()
                            with a2.p():
                                a2.t(style="font-weight:bold", _t='Contact us:')
                            with a2.p():
                                a2('If you experience any difficulties using this item download or you have any questions, please contact us at Ask a Schlesinger Librarian: asklib.schlesinger.radcliffe.edu')
                            a2.br()
                            with a2.div(klass="page2 pb_before"):
                                a2.br()
                                with a2.p():
                                    a2("The following files are navigable using the bookmark tools in your PDF viewer. For example, in Adobe Acrobat Pro, use the ribbon icon on the left panel to view and navigate each file.")
                                a2.br()
                                with a2.p():
                                    a2.t(style="font-weight:bold", _t='Files in this item:<br/>')
                                    a2('<br/>'.join(item_names))
        html2 = str(a2)
        # print(html2)
        stylesheet_path = "C:\Temp\PrependPDF-css\\"
        css = stylesheet_path + 'stylesheet.css'
        pdfkit.from_string(html2, folder_path +'\\'+ E_number +'_cover'+ '.pdf', css=css)

    # define pre-merge folder
    pre_merge_folder = folder_path + "\\" + "pre-merge"
    vrr_delivery_folder = folder_path + "\\" + "vrr_delivery"
    if not VRR_check:
        if not os.path.exists(pre_merge_folder):
            os.makedirs(pre_merge_folder)

    # parse finding aid
    for file in os.listdir(folder_path):
        if file.startswith("collection"):
            collection_name = '\\'+file
            collection_path = folder_path+collection_name
            print('Examining ' + file)
    try:
        collection_path
    except NameError:
        print("ERROR: Finding aid not found. Download a finding aid from the PUI and add it to the source directory.")
        continue
    Collection_data = []
    E_number = input('Enter the E number assigned to this archival item. For example, "E.2"\n')

    with open(collection_path, encoding="utf-8-sig") as FindingAid_obj:
        CollectionReader_obj = csv.reader(FindingAid_obj)
        for row in CollectionReader_obj:
            Collection_data.append(row)
        for row in Collection_data:
            if "Collection Title" in row:
                Collection_name = row[1]
        for row in Collection_data:
            if "Collection Dates" in row:
                Collection_dates = row[1]
        for row in Collection_data:
            if "Collection Identifier" in row:
                Full_ID = row[1]
        E_number_search_string = E_number + '.'
        print(E_number_search_string)
        Collection_ID_0 = Full_ID.replace(':', ';')
        Collection_ID = Collection_ID_0.split(';', 1)[0]
        Collection_ID2 = Collection_ID.replace(' ', '')
        for row in Collection_data:
            if "EAD ID" in row:
                EAD_ID = row[1]
        Permalink = "id.lib.harvard.edu/ead/" + EAD_ID + "/catalog"
        try:
            for row in Collection_data:
                if E_number_search_string in row[5]:
                    Database_number = row[0]
                    Folder_name = row[1]
                    if len(Folder_name) > 90:
                        Folder_truncated = (Folder_name[:90] + '...')
                        print("Truncating to " + Folder_truncated)
                        Folder_name = Folder_truncated
                    else:
                        print("Folder name is " + Folder_name)
        except NameError:
            continue
        try:
            print('"' + Folder_name + '"' + " found in FA.")
        except NameError:
            input('"' + E_number_search_string + '"' + " not found in FA file. Check that the search string exists in Column F and has a trailing period then restart PrependPDF.")

    def paginate_PDF():
        input_item_path2 = pre_merge_folder + "//" + Item_number + "-letter.pdf"
        src = fitz.open(input_item_path)
        doc = fitz.open()
        for ipage in src:
            irotation = ipage.rotation
            iwidth = ipage.rect.width
            iheight = ipage.rect.height
            if ipage.rotation !=0:
                ipage.set_rotation(int(float(ipage.rotation))-(int(float(ipage.rotation))))
                iwidth, iheight = iheight, iwidth
            # if iwidth > iheight:
            #     fmt = fitz.paper_rect("letter-l")
            # else:
            fmt = fitz.paper_rect("letter")
            page = doc.new_page(width=fmt.width, height=fmt.height)
            page.show_pdf_page(page.rect, src, ipage.number, rotate=-int(float(irotation)))

        doc.save(input_item_path2)

        pdfmetrics.registerFont(TTFont('Inter', 'C:\Temp\PrependPDF-css\Inter.ttf'))
        input_file = input_item_path2
        output_file = item_path

        # Get pages
        reader = PdfReader(input_file)


        pages = [pagexobj(p) for p in reader.pages]
        canvas = Canvas(output_file, pagesize=letter)


        # Compose new pdf

        for page_num, page in enumerate(pages, start=1):
            canvas.setPageSize((612, 792))
            canvas.scale(0.88, 0.88)
            canvas.translate(39.685, 48.189)
            canvas.doForm(makerl(canvas, page))

            # Draw footer
            schlesinger_footer_text = "Schlesinger Library • Harvard Radcliffe Institute"
            collection_footer_text = Collection_ID2 + " • " + Collection_name
            try:
                folder_header_text = E_number + " • " + Folder_name
            except NameError:
                input('"' + E_number_search_string + '"' + " not found in FA file. Check that the search string exists in Column F and has a trailing period then restart PrependPDF.")
            item_header_text = Name
            page_header_text = "Page %s of %s" % (page_num, len(pages))
            canvas.saveState()
            canvas.setStrokeColorRGB(0, 0, 0)
            canvas.setLineWidth(0.75)
            canvas.line(25, -2, 612 - 25, -2)
            canvas.line(25, 792+2, 612 - 25, 792+2)
            canvas.setFont('Inter', 10)
            canvas.drawString(25, 792+22, item_header_text)
            canvas.drawString(25, 792+9, folder_header_text)
            canvas.drawString(25, -16, schlesinger_footer_text)
            canvas.drawString(25, -29, collection_footer_text)
            canvas.drawRightString(612-25, 792+22, page_header_text)
            canvas.restoreState()

            canvas.showPage()

        canvas.save()

    def readme_writer():
        if VRR_check is True:
            print("VRR_check is True.")
            readme_path = vrr_delivery_folder + '\\' + 'readme_' + E_number + '.txt'
            lines = [
                'Readme',
                '',
                'This directory contains access versions of files from the',
                'following collection and archival item:',
                '',
                Collection_ID + ". " + Collection_name,
                E_number + ". " + Folder_name,
                '',
                'To return to the finding aid for this collection, please use the',
                'following link:',
                'https://id.lib.harvard.edu/ead/c/' + Database_number + '/catalog',
                '',
                'The files stored in this directory are presented as transformed',
                'versions with their original formatting preserved as closely as',
                'possible. Where necessary, files have been converted so they may',
                'be accessed using widely available software such as Microsoft',
                'Office, LibreOffice, etc.',
                '',
                'The following files are stored and named using unique identifiers:',
                '\n',
            ]
        elif VRR_PDF_check is True:
            print("VRR_PDF_check is True.")
            readme_path = product_folder + '\\' + 'readme_' + E_number + '.txt'
            lines = [
                'Readme',
                '',
                'This directory contains access versions of files from the',
                'following collection:',
                '',
                Collection_ID + ". " + Collection_name,
                E_number + ". " + Folder_name,
                '',
                'The files stored in this directory are presented as transformed',
                'versions with their original formatting preserved as closely as',
                'possible. Files have been converted to PDF for easier distribution',
                'and access.',

                'The following archival items are stored and named using unique',
                'identifiers:',
                '',
                '',
                '***REPLACE WITH LIST OF ARCHIVAL ITEMS***'
                '\n',
            ]
        lines.extend(vrr_filelist)
        with open(readme_path, 'w') as f:
            f.write('\n'.join(lines))


    # create lists for merging
    dir_list = [f for f in listdir(folder_path) if isfile(join(folder_path, f))]
    file_list = []
    covers = []
    covers2 = []
    items = []
    item_names = []
    bookmark_list = []
    vrr_filelist = []

    # use file list data and finding aid to generate item covers and paginate items
    if VRR_check is not True:
        print('Generating item cover pages...')

    #checking for tsv
    if os.path.isfile(folder_path+'\FileList.tsv'):
        tsv_file = folder_path+'\FileList.tsv'
        csv_table = pd.read_table(tsv_file, sep='\t')
        csv_table.to_csv(folder_path+'\FileList.csv', index=False)
        print("TSV converted to CSV for reading.")
    elif os.path.isfile(folder_path+'\FileList.txt'):
        txt_file = folder_path+'\FileList.txt'
        csv_table = pd.read_table(txt_file, sep='\t')
        csv_table.to_csv(folder_path+'\FileList.csv', index=False)
        print("TSV converted to CSV for reading.")
    if os.path.isfile(folder_path+'\FileList.csv'):
        with open(folder_path+'\FileList.csv', encoding="utf-8-sig") as FileList_obj:
            FileReader_obj = csv.DictReader(FileList_obj, dialect=csv.excel)
            for row in FileReader_obj:
                Name = str(row["Name"])
                Item_number = str(row['Item #'])
                try:
                    Size = str(round(int(row['P-Size (bytes)']) / 1000, 1)) + ' Kb'
                except:
                    Size = str(round(int(row['L-Size (bytes)']) / 1000, 1)) + ' Kb'
                if not str(row['Category']):
                    Category = "Not available"
                else:
                    Category = str(row['Category'])
                if not str(row['Created']):
                    Created_date = "Not available"
                elif str(row['Created']) == "n/a":
                    Created_date = "Not available"
                else:
                    split_Created_date = row['Created'].split("(")[1]
                    Created_date = calendar.month_name[(int(split_Created_date[5:7]))] + " " + split_Created_date[8:10].lstrip("0") + ", " + split_Created_date[:4] + " at " + split_Created_date[10:16]
                if not str(row["Modified"]):
                    Modified_date = "Not available"
                elif str(row['Modified']) == "n/a":
                    Modified_date = "Not available"
                else:
                    split_Modified_date = row['Modified'].split("(")[1]
                    Modified_date = calendar.month_name[(int(split_Modified_date[5:7]))] + " " + split_Modified_date[8:10].lstrip("0") + ", " + split_Modified_date[:4] + " at " + split_Modified_date[10:16]
                if not str(row['Author']):
                    Author = 'Not available'
                else:
                    Author = str(row['Author'])
                if not str(row['Last Saved By']):
                    Last_saved_by = 'Not available'
                else:
                    Last_saved_by = str(row['Last Saved By'])
                for f in dir_list:
                    if f[:len(Item_number)] == Item_number:
                        input_item_path = folder_path + "\\" + f
                        file_list.append(Item_number)
                item_names.append(Name)
                item_path = pre_merge_folder + "\\" + Item_number + '.pdf'
                if VRR_check is True:
                    vrr_filelist.append("File:   " + Item_number)
                    vrr_filelist.append("   Original file name:       " + Name)
                    vrr_filelist.append("   Original file category:   " + Category)
                    vrr_filelist.append("   File size on disk:        " + Size)
                    vrr_filelist.append("   Date created:             " + Created_date)
                    vrr_filelist.append("   Last modified:            " + Modified_date)
                    vrr_filelist.append("   Author:                   " + Author)
                    vrr_filelist.append("   Last saved by:            " + Last_saved_by)
                    vrr_filelist.append("\n")
                    if not os.path.exists(vrr_delivery_folder):
                        os.makedirs(vrr_delivery_folder)
                    if os.path.exists(vrr_delivery_folder):
                        shutil.copy(input_item_path, vrr_delivery_folder)
                else:
                    cover_name = Item_number + '_cover' + '.pdf'
                    covers.append(cover_name)
                    bookmark_list.append(Item_number)
                    bookmark_list.append(Name)
                    covers2.append(pre_merge_folder + "\\" + cover_name)
                    items.append(item_path)
                    if (input_item_path[-4:] == '.jpg' or input_item_path[-4:] == '.png' or input_item_path[-4:] == '.tif' or input_item_path[-4:] == '.jp2' ):
                        input_image = input_item_path
                        image_PDF_path = input_image[:-4] + '.pdf'
                        a4inpt = (img2pdf.mm_to_pt(215.9), img2pdf.mm_to_pt(279.4))
                        layout_fun = img2pdf.get_layout_fun(a4inpt)
                        image_object = open(image_PDF_path, "wb")
                        with image_object as f:
                            f.write(img2pdf.convert(input_image, layout_fun=layout_fun))
                        input_item_path = image_PDF_path
                    if VRR_check is not True:
                        file_pages = open(input_item_path, 'rb')
                        # try:
                        try:
                            pdfReader = PyPDF2.PdfReader(file_pages)
                        except:
                            print(
                                'ERROR 1: The application failed to read one or more PDFs.')
                            continue
                        Page_count = str(len(pdfReader.pages)) + ' pages'
                    else:
                        Page_count = 'Not available'
                    create_item_cover()
                    if not os.path.exists(pre_merge_folder):
                        os.makedirs(pre_merge_folder)
                    if os.path.exists(pre_merge_folder):
                        cover_path = folder_path + "\\" + cover_name
                        if os.path.isfile(pre_merge_folder + '\\' + cover_name):
                            os.remove(pre_merge_folder + '\\' + cover_name)
                        shutil.move(cover_path, pre_merge_folder)

                        paginate_PDF()
                        # except:
                        #     print('ERROR 2: The application failed to read one or more PDFs.')
                    print('Covered '+Item_number+'.')

    if VRR_check is not True:
        try:
            for idx, item in enumerate(items):
                item_count = str(idx + 1)
            print('Covered ' + item_count + ' items.')
        except:
            print("Unable to count items.")

    elif not os.path.exists(folder_path+'\FileList.csv'):
        print("ERROR: FileList.csv not found.")
        continue

    else:
        print("VRR files packaged.")

    product_folder = folder_path + "\\" + "product"

    if not os.path.exists(product_folder):
        os.makedirs(product_folder)

    if VRR_check is True or VRR_PDF_check is True:
       readme_writer()

    # add VRR delivery column to FileList.csv

    new_VRR_column = []

    for i in file_list:
        if VRR_check is True:
            new_VRR_column.append('y')
        elif VRR_PDF_check is True:
            new_VRR_column.append('y')
        else:
            new_VRR_column.append('')
    editFileList = pd.read_csv(folder_path+'\FileList.csv')
    new_column = pd.DataFrame({'VRR_delivery': new_VRR_column})
    editFileList['EADID'] = EAD_ID
    editFileList['ComponentID'] = Database_number
    editFileList['MC_number'] = Collection_ID2
    editFileList['Full_E_number'] = Database_number + '--' + Collection_ID2 + '_' + E_number
    editFileList['E_number'] = E_number
    editFileList['BDID'] = editFileList['Path'].astype(str).str[:8]
    editFileList = editFileList.merge(new_column, left_index = True, right_index = True)
    editFileList.to_csv(product_folder+'\FileList' + '_' + Database_number + '--' + Collection_ID2 + '_' + E_number + '.csv', index = False)
    print('FileList.csv ready for import to BDT.')

    if VRR_check is not True:
        # use finding aid to generate the cover page for the E.#
        print('Generating ' + E_number + ' cover page...')
        create_master_cover()
        master_cover_name = E_number +'_cover'+ '.pdf'
        master_cover_path = folder_path +'\\'+ master_cover_name
        if os.path.isfile(pre_merge_folder + '\\' + master_cover_name):
            os.remove(pre_merge_folder + '\\' + master_cover_name)
        shutil.move(master_cover_path, pre_merge_folder)
        print(E_number + ' covered.')

        # gather populated lists into a master list to parse pre-merge
        print('Assembling PDFs...')
        merged_pdfs = [None]*(len(covers2)+len(items))
        merged_pdfs[::2] = covers2
        merged_pdfs[1::2] = items
        merged_pdfs.insert(0, pre_merge_folder + '\\' + master_cover_name)

        # print(merged_pdfs)
        print('Merging PDFs...')

        # merge PDFs

        merger = PdfMerger()
        metadata = {
            u'/Version': 'PDF-1.6'
        }
        try:
            for pdf in merged_pdfs:
                if pdf in covers2:
                    base_name = os.path.basename(pdf)
                    base_ID = base_name.split('_cover')[0]
                    if base_ID not in bookmark_list:
                        bookmark = None
                    else:
                        bookmark = bookmark_list[bookmark_list.index(base_ID)+1]
                    bookmark = bookmark
                elif pdf is merged_pdfs[0]:
                    bookmark = E_number + ' cover page'
                else:
                    bookmark = None
                merger.append(pdf, bookmark)
                merger.add_metadata(metadata)
            source_PDF_path = pre_merge_folder + '\\' + Database_number + '--' + Collection_ID2 + '_' + E_number + '_SOURCE.pdf'

            merger.write(source_PDF_path)
            merger.close()

            print('Converting to PDF/A...')

            product_path = product_folder + '\\' + Database_number + '--' + Collection_ID2 + '_' + E_number + '.pdf'

            import ghostscript

            args = [
                "pdf2pdfa",
                "--permit-file-read=C:\\Windows\\System32\\spool\\drivers\\color\\AdobeRGB1998.icc",
                "-dPDFA=2", "-dBATCH", "-dNOPAUSE", "-dSAFER",
                "-sColorConversionStrategy=RGB",
                "-sDEVICE=pdfwrite", "-dPDFACompatibilityPolicy=1",
                "-sProfileOut=" + "C:\Windows\System32\spool\drivers\color\AdobeRGB1998.icc",
                "-sOutputFile=" + product_path,
                "C:\\Program Files\\gs\\gs10.00.0\\lib\\PDFA_def.ps",
                source_PDF_path
            ]

            ghostscript.Ghostscript(*args)

            print('PDFs merged and saved as ' + Database_number + '--' + Collection_ID2 + '_' + E_number + '.pdf. \n')
        except:
            print('ERROR 3: The application failed to read one or more PDFs.')
            continue