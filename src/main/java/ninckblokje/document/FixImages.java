package ninckblokje.document;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

import static javax.xml.xpath.XPathConstants.NODESET;

public class FixImages {

    public static void main(String[] args) throws Exception {
        Document doc = getDocumentBuilderFactory().newDocumentBuilder().parse(new File("data/working/word/document.xml"));

        XPath xPath = XPathFactory.newInstance().newXPath();
        xPath.setNamespaceContext(getNamespaceContext());

        XPathExpression xPathExpression = xPath.compile("//a:blip");
        NodeList nodeList = (NodeList) xPathExpression.evaluate(doc.getDocumentElement(), NODESET);
        System.out.println(String.format("%d blip nodes found", nodeList.getLength()));

        List<String> embeddedRelationIds = new ArrayList<>();
        for (int i = 0; i < nodeList.getLength(); i++) {
            Node node = nodeList.item(i);
            Node embed = node.getAttributes().getNamedItemNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
            embeddedRelationIds.add(embed.getTextContent());
        }
        System.out.println(String.format("%d embedded relation id found", embeddedRelationIds.size()));

        List<String> mediaFiles = parseMediaFiles();
        Map<String, String> relations = parseRelations();

        List<String> foundMediaFiles = new ArrayList<>();
        List<String> notFoundMediaFiles = new ArrayList<>();

        Collection<String> relationValues = relations.values();
        mediaFiles.forEach(mediaFile -> {
            if (relationValues.contains(mediaFile)) {
                foundMediaFiles.add(mediaFile);
            } else {
                notFoundMediaFiles.add(mediaFile);
            }
        });
        System.out.println(String.format("%d found media files, %d not found media files", foundMediaFiles.size(), notFoundMediaFiles.size()));

        List<String> missingMediaFiles = new ArrayList<>();

        relations.forEach((key, value) -> {
            if (!value.toLowerCase().endsWith(".xml") && !mediaFiles.contains(value)) {
                missingMediaFiles.add(value);
            }
        });
        System.out.println(String.format("%d missing media files", missingMediaFiles.size()));

        List<String> foundEmbeddedRelationIds = new ArrayList<>();
        List<String> notFoundEmbeddedRelationIds = new ArrayList<>();

        embeddedRelationIds.stream().forEach(embeddedRelationId -> {
            if (relations.containsKey(embeddedRelationId)) {
                foundEmbeddedRelationIds.add(embeddedRelationId);
            } else {
                notFoundEmbeddedRelationIds.add(embeddedRelationId);
            }
        });
        System.out.println(String.format("%d found embedded relation ids, %d not found embedded relation ids", foundEmbeddedRelationIds.size(), notFoundEmbeddedRelationIds.size()));

        System.err.println(String.format(
                "Erroneous embedded relation ids: %s",
                notFoundEmbeddedRelationIds.stream().distinct().collect(Collectors.toList())
        ));
    }

    static DocumentBuilderFactory getDocumentBuilderFactory() {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newDefaultInstance();
        dbf.setNamespaceAware(true);
        return dbf;
    }

    static NamespaceContext getNamespaceContext() {
        return new NamespaceContext() {

            @Override
            public String getNamespaceURI(String prefix) {
                switch (prefix) {
                    case "a":
                        return "http://schemas.openxmlformats.org/drawingml/2006/main";
                    case "pr":
                        return "http://schemas.openxmlformats.org/package/2006/relationships";
                    case "r":
                        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    default:
                        throw new IllegalArgumentException(String.format("Unknown namespace prefix %s", prefix));
                }
            }

            @Override
            public String getPrefix(String namespaceURI) {
                switch (namespaceURI) {
                    case "http://schemas.openxmlformats.org/drawingml/2006/main":
                        return "a";
                    case "http://schemas.openxmlformats.org/package/2006/relationships":
                        return "pr";
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships":
                        return "r";
                    default:
                        throw new IllegalArgumentException(String.format("Unknown namespace URI %s", namespaceURI));
                }
            }

            @Override
            public Iterator<String> getPrefixes(String namespaceURI) {
                switch (namespaceURI) {
                    case "http://schemas.openxmlformats.org/drawingml/2006/main":
                        return Arrays.asList("a").iterator();
                    case "http://schemas.openxmlformats.org/package/2006/relationships":
                        return Arrays.asList("pi").iterator();
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships":
                        return Arrays.asList("r").iterator();
                    default:
                        throw new IllegalArgumentException(String.format("Unknown namespace URI %s", namespaceURI));
                }
            }
        };
    }

    static List<String> parseMediaFiles() {
        List<String> mediaFiles = new ArrayList<>();

        Arrays.stream(new File("data/working/word/media").listFiles())
                .forEach(file -> mediaFiles.add("media/" + file.getName()));

        System.out.println(String.format("%d media files found", mediaFiles.size()));
        return mediaFiles;
    }

    static Map<String, String> parseRelations() throws ParserConfigurationException, XPathExpressionException, IOException, SAXException {
        Document relDoc = getDocumentBuilderFactory().newDocumentBuilder().parse(new File("data/working/word/_rels/document.xml.rels"));
        XPath relXpath = XPathFactory.newInstance().newXPath();
        relXpath.setNamespaceContext(getNamespaceContext());

        XPathExpression relXPathExpression = relXpath.compile("//pr:Relationship");
        NodeList relNodeList = (NodeList) relXPathExpression.evaluate(relDoc.getDocumentElement(), NODESET);

        System.out.println(String.format("%d relation nodes found", relNodeList.getLength()));

        Map<String, String> relations = new HashMap<>();

        for (int i = 0; i < relNodeList.getLength(); i++) {
            Node relNode = relNodeList.item(i);
            Node idNode = relNode.getAttributes().getNamedItem("Id");
            Node targetNode = relNode.getAttributes().getNamedItem("Target");

            if (targetNode.getNodeValue().toUpperCase().equals("NULL")) {
                System.err.println(idNode.getTextContent() + " has a NULL target");
            } else {
                relations.put(idNode.getTextContent(), targetNode.getTextContent());
            }
        }

        return relations;
    }
}
