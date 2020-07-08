package ninckblokje.document;

import org.apache.commons.lang3.tuple.ImmutableTriple;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import static javax.xml.xpath.XPathConstants.*;

public class FixImages {

    public static void main(String[] args) throws Exception {
        Document doc = getDocumentBuilderFactory().newDocumentBuilder().parse(new File("data/working/word/document.xml"));
        Document relDoc = getDocumentBuilderFactory().newDocumentBuilder().parse(new File("data/working/word/_rels/document.xml.rels"));

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
        Map<String, String> relations = parseRelations(relDoc);

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

        List<Graphic> graphics = parseGraphics(doc);
        graphics.stream()
                .filter(graphic -> graphic.getRelationId().equals("rId7"))
                .forEach(graphic -> System.out.println(String.format("%d - %s - %s", graphic.getIndex() + 1, graphic.getRelationId(), graphic.getFilename())));

//        ImageFile maxImageFile = getMaxImageFile();
//        RId maxRId = getMaxRId(relDoc);
//        addMissingImages(doc, relDoc, maxRId, maxImageFile);
//
//        saveDocument(relDoc, new File("data/working/word/_rels/document.xml.rels"));
//        saveDocument(doc, new File("data/working/word/document.xml"));
    }

    static void addMissingImages(Document doc, Document relDoc, RId rId, ImageFile imageFile) throws IOException {
        ImageFileWalker imageFileWalker = new ImageFileWalker(imageFile);
        RIdWalker rIdWalker = new RIdWalker(rId);

        Files.list(Path.of("data/missingImages"))
                .filter(path -> isImageMissing(doc, path))
                .map(path -> ImmutableTriple.of(path, rIdWalker, imageFileWalker))
                .forEach(pair -> addMissingImage(doc, relDoc, pair.getMiddle(), pair.getRight(), pair.getLeft()));
    }

    static void addMissingImage(Document doc, Document relDoc, RIdWalker rIdWalker, ImageFileWalker imageFileWalker, Path path) {
        double missingImageCount = imageMissingCount(doc, path);

        System.out.println(path.getFileName() + " is missing " + missingImageCount + " times");

        try {
            for (double i=0.0; i<missingImageCount; i=i+1.0) {
                RId rId = rIdWalker.next();
                ImageFile imageFile = imageFileWalker.next();

                System.out.println(path.getFileName() + " will get " + rId.getrId() + " on index " + i);

                Path targetFile = Path.of("data/working/word/media", imageFile.getFilename());
                System.out.println("Copying " + path + " to " + targetFile);
                Files.copy(path, targetFile);

                createNewRelationship(relDoc, rId, imageFile);
                updateGraphics(doc, rId, path);
            }

            Path doneFile = Path.of("data/done", path.getFileName().toString());
            System.out.println("Done with " + path);
            Files.move(path, doneFile);
        } catch (IOException | XPathExpressionException ex) {
            throw new RuntimeException(ex);
        }
    }

    static void updateGraphics(Document doc, RId rId, Path path) throws XPathExpressionException {
        XPath xPath = XPathFactory.newInstance().newXPath();
        xPath.setNamespaceContext(getNamespaceContext());

        XPathExpression idXPathExpression = xPath.compile(String.format("//a:graphic/a:graphicData/pic:pic[pic:nvPicPr/pic:cNvPr/@name=\"%s\"]/pic:blipFill/a:blip[@r:embed=\"rId7\"]", path.getFileName()));
        Element blibElement = (Element) idXPathExpression.evaluate(doc, NODE);
        blibElement.setAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed", rId.getrId());
    }

    static void createNewRelationship(Document relDoc, RId rId, ImageFile imageFile) {
        System.out.println("Creating new relationship " + rId.getrId() + " " + imageFile.getFilename());
        Element newRelationship = relDoc.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
        newRelationship.setAttribute("Id", rId.getrId());
        newRelationship.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        newRelationship.setAttribute("Target", String.format("media/%s", imageFile.getFilename()));
        relDoc.getDocumentElement().appendChild(newRelationship);
    }

    static DocumentBuilderFactory getDocumentBuilderFactory() {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newDefaultInstance();
        dbf.setNamespaceAware(true);
        return dbf;
    }

    static ImageFile getMaxImageFile() throws IOException {
        List<ImageFile> imagesFiles = Files.list(Path.of("data/working/word/media"))
                .map(Path::getFileName)
                .map(Path::toString)
                .map(ImageFile::new)
                .collect(Collectors.toList());
        System.out.println("Finding max image file over " + imagesFiles.size() + " images ");

        ImageFile maxImageFile = null;
        for(ImageFile imageFile: imagesFiles) {
            if (maxImageFile == null) {
                maxImageFile = imageFile;
            } else if (imageFile.isLargerThen(maxImageFile)) {
                maxImageFile = imageFile;
            }
        }

        System.out.println("Found " + maxImageFile);
        return maxImageFile;
    }

    static RId getMaxRId(Document relDoc) throws XPathExpressionException {
        XPath xPath = XPathFactory.newInstance().newXPath();
        xPath.setNamespaceContext(getNamespaceContext());

        XPathExpression idXPathExpression = xPath.compile("//pr:Relationship/@Id");
        NodeList nodeSet = (NodeList) idXPathExpression.evaluate(relDoc, NODESET);
        System.out.println("Finding max rId over " + nodeSet.getLength() + " relation nodes");

        RId rId = null;
        for (int i = 0; i < nodeSet.getLength(); i++) {
            RId tempRId = new RId(nodeSet.item(i).getTextContent());
            if (rId == null || tempRId.isLargerThen(rId)) {
                rId = tempRId;
            }
        }

        System.out.println("Found " + rId);
        return rId;
    }

    static boolean isImageMissing(Document doc, Path path) {
        return imageMissingCount(doc, path) >= 1.0;
    }

    static double imageMissingCount(Document doc, Path path) {
        XPath xPath = XPathFactory.newInstance().newXPath();
        xPath.setNamespaceContext(getNamespaceContext());

        XPathExpression countXPathExpression = null;
        try {
            countXPathExpression = xPath.compile(String.format("count(//a:graphic/a:graphicData/pic:pic[pic:nvPicPr/pic:cNvPr/@name=\"%s\"]/pic:blipFill/a:blip[@r:embed=\"rId7\"])", path.getFileName()));
            double count = (Double) countXPathExpression.evaluate(doc, NUMBER);

            System.out.println(path.getFileName() + " is missing: " + count);
            return count;
        } catch (XPathExpressionException ex) {
            throw new RuntimeException(ex);
        }
    }

    static NamespaceContext getNamespaceContext() {
        return new NamespaceContext() {

            @Override
            public String getNamespaceURI(String prefix) {
                switch (prefix) {
                    case "a":
                        return "http://schemas.openxmlformats.org/drawingml/2006/main";
                    case "pic":
                        return "http://schemas.openxmlformats.org/drawingml/2006/picture";
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
                    case "http://schemas.openxmlformats.org/drawingml/2006/picture":
                        return "pic";
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
                    case "http://schemas.openxmlformats.org/drawingml/2006/picture":
                        return Arrays.asList("pic").iterator();
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

    static List<Graphic> parseGraphics(Document doc) throws XPathExpressionException {
        XPath xPath = XPathFactory.newInstance().newXPath();
        xPath.setNamespaceContext(getNamespaceContext());

        XPathExpression graphicXPathExpression = xPath.compile("//a:graphic");
        XPathExpression pictureXPathExpression = xPath.compile(".//pic:cNvPr");
        XPathExpression blipXPathExpression = xPath.compile(".//a:blip");

        NodeList graphicNodeList = (NodeList) graphicXPathExpression.evaluate(doc.getDocumentElement(), NODESET);

        System.out.println(String.format("%d graphics found", graphicNodeList.getLength()));

        List<Graphic> graphics = new ArrayList<>();

        for (int i = 0; i < graphicNodeList.getLength(); i++) {
            Node graphicNode = graphicNodeList.item(i);

            Node pictureNode = (Node) pictureXPathExpression.evaluate(graphicNode, NODE);
            Node blipNode = (Node) blipXPathExpression.evaluate(graphicNode, NODE);
            if (pictureNode != null && blipNode != null) {
                Node nameNode = pictureNode.getAttributes().getNamedItem("name");
                Node embedNode = blipNode.getAttributes().getNamedItemNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");

                graphics.add(new Graphic(i, embedNode.getTextContent(), nameNode.getTextContent()));
            }
        }

        return graphics;
    }

    static List<String> parseMediaFiles() {
        List<String> mediaFiles = new ArrayList<>();

        Arrays.stream(new File("data/working/word/media").listFiles())
                .forEach(file -> mediaFiles.add("media/" + file.getName()));

        System.out.println(String.format("%d media files found", mediaFiles.size()));
        return mediaFiles;
    }

    static List<String> parsePictures(Document doc) throws XPathExpressionException {
        XPath pictureXPath = XPathFactory.newInstance().newXPath();
        pictureXPath.setNamespaceContext(getNamespaceContext());

        XPathExpression pictureXPathExpression = pictureXPath.compile("//pic:cNvPr");
        NodeList pictureNodeList = (NodeList) pictureXPathExpression.evaluate(doc.getDocumentElement(), NODESET);

        System.out.println(String.format("%d pictures found", pictureNodeList.getLength()));

        List<String> pictures = new ArrayList<>();

        for (int i = 0; i < pictureNodeList.getLength(); i++) {
            Node pictureNode = pictureNodeList.item(i);
            Node nameNode = pictureNode.getAttributes().getNamedItem("name");

            pictures.add(nameNode.getTextContent());
        }

        return pictures;
    }

    static Map<String, String> parseRelations(Document relDoc) throws ParserConfigurationException, XPathExpressionException, IOException, SAXException {
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

    static int parseImageFilename(String filename) {
        Pattern imageFilenamePattern = Pattern.compile("^image(\\d+)\\.");
        Matcher m = imageFilenamePattern.matcher(filename);

        m.find();

        return Integer.parseInt(m.group(1));
    }

    static int parseRIdValue(String rId) {
        Pattern rIdPattern = Pattern.compile("^(rId)(\\d+)$");
        Matcher m = rIdPattern.matcher(rId);

        m.find();

        return Integer.parseInt(m.group(2));
    }

    static void saveDocument(Document doc, File xmlFile) throws TransformerException, IOException {
        System.out.println("Storing document " + xmlFile);
        TransformerFactory tf = TransformerFactory.newDefaultInstance();
        tf.newTransformer().transform(new DOMSource(doc), new StreamResult(new FileWriter(xmlFile)));
    }
}

class Graphic {

    private int index;
    private String relationId;
    private String filename;

    public Graphic(int index, String relationId, String filename) {
        this.index = index;
        this.relationId = relationId;
        this.filename = filename;
    }

    public int getIndex() {
        return index;
    }

    public String getRelationId() {
        return relationId;
    }

    public String getFilename() {
        return filename;
    }
}

class ImageFile {

    private String filename;
    private int idValue;

    public ImageFile(String filename) {
        this.filename = filename;
        this.idValue = FixImages.parseImageFilename(filename);
    }

    public ImageFile(int idValue) {
        this.filename = String.format("image%d.jpeg", idValue);
        this.idValue = idValue;
    }

    public String getFilename() {
        return filename;
    }

    public int getIdValue() {
        return idValue;
    }

    public boolean isLargerThen(ImageFile otherImageFile) {
        return this.idValue > otherImageFile.getIdValue();
    }

    @Override
    public String toString() {
        return String.format("%s %d", filename, idValue);
    }

    public static ImageFile next(ImageFile imageFile) {
        return new ImageFile(imageFile.getIdValue() + 1);
    }
}

class ImageFileWalker {

    private ImageFile imageFile;

    public ImageFileWalker(ImageFile imageFile) {
        this.imageFile = imageFile;
    }

    public ImageFile current() {
        return imageFile;
    }

    public ImageFile next() {
        imageFile = ImageFile.next(imageFile);
        return imageFile;
    }
}

class RId {

    private final String rId;
    private final int idValue;

    public RId(String rId) {
        this.rId = rId;
        this.idValue = FixImages.parseRIdValue(this.rId);
    }

    public RId(int idValue) {
        this.rId = String.format("rId%d", idValue);
        this.idValue = idValue;
    }

    public String getrId() {
        return rId;
    }

    public int getIdValue() {
        return idValue;
    }

    public boolean isLargerThen(RId otherRId) {
        return this.idValue > otherRId.getIdValue();
    }

    @Override
    public String toString() {
        return String.format("%s %d", rId, idValue);
    }

    public static RId next(RId rId) {
        return new RId(rId.getIdValue() + 1);
    }
}

class RIdWalker {

    private RId rId;

    public RIdWalker(RId rId) {
        this.rId = rId;
    }

    public RId current() {
        return rId;
    }

    public RId next() {
        rId = RId.next(rId);
        return rId;
    }
}
