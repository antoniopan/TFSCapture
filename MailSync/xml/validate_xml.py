import lxml.etree as et


def validate_xml(xsd_path, xml_path):
    print("Validate '%s' according to '%s'..." % (xml_path, xsd_path))
    xml_doc = et.parse(xml_path)
    xmlschema_doc = et.parse(xsd_path)
    xmlschema = et.XMLSchema(xmlschema_doc)
    if True == xmlschema.validate(xml_doc):
        print("Validation suceeded.")
    else:
        print("Validation failed.")
        print(xmlschema.error_log)


if __name__ == '__main__':
    validate_xml('E:/Code/CSharp/TFSCapture/TfsTracker/ProjectSchema.xsd', 'E:/Code/CSharp/TFSCapture/MailSync/xml/dsa_query.xml')
    validate_xml('E:/Code/CSharp/TFSCapture/TfsTracker/ProjectSchema.xsd', 'E:/Code/CSharp/TFSCapture/MailSync/xml/uideal_query.xml')