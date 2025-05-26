import xml.etree.ElementTree as ET

def analyze_xml(xml_file):
    """
    Analyzes an XML file to extract all titles and count specific tags.

    Args:
        xml_file (str): Path to the XML file.

    Returns:
        tuple: (titles, media_object_count, figure_count)
    """
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Extract all titles
        title_elements = root.findall('.//{http://xml.cengage-learning.com/cendoc-core}title')
        titles = [title.text if title.text is not None else "Title not found" for title in title_elements]

        # Count media-object tags
        media_object_count = len(root.findall('.//{http://xml.cengage-learning.com/cendoc-core}media-object'))

        # Count figure tags
        figure_count = len(root.findall('.//{http://xml.cengage-learning.com/cendoc-core}figure'))

        return titles, media_object_count, figure_count

    except FileNotFoundError:
        return ["File not found"], 0, 0
    except ET.ParseError:
        return ["XML Parse Error"], 0, 0

if __name__ == "__main__":
    xml_file = "smaple_data/cendoc-979118-PPEJLQPGBBJDMYAB0158.xml"
    titles, media_object_count, figure_count = analyze_xml(xml_file)

    print("Titles:")
    for title in titles:
        print(f"- {title}")
    print(f"Media Object Count: {media_object_count}")
    print(f"Figure Count: {figure_count}")
