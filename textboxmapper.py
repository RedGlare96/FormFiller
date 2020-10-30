import pdfrw

if __name__ == '__main__':
    filename = input('Enter the path to pdf file to be analyzed: \n')
    template = pdfrw.PdfReader(filename)
    with open('{}-mapped.txt'.format(filename.split('.')[0]), 'w') as f:
        for page in range(len(template.pages)):
            annotations = template.pages[page]['/Annots']
            for item in annotations:
                if item['/Subtype'] == '/Widget':
                    if item['/T']:
                        key = item['/T'][1: -1]
                        f.write(key + '\n')
                        item.update(pdfrw.PdfDict(V=key))
    outputfile = '{}-mapped.pdf'.format(filename.split('.')[0])
    print('Writing: {}'.format(outputfile))
    pdfrw.PdfWriter().write(outputfile, template)