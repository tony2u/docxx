#include "docxx.hpp"
#include "utils.hpp"
#include <io.h>

#include <iostream>
#include <random>

using namespace docxx;

// Document class
Document::Document(const std::string&& file, bool readOnce = false) : doc_file(file) {
    if (readOnce) this->open();
}
Document::Document(const char* file, bool readOnce = false) : doc_file(file) {
    if (readOnce) this->open();
}
pugi::xml_node& Document::get_body() const {
    return this->_body;
}
pugi::xml_document& Document::get_document() const {
    return this->_document;
}
pugi::xml_document& Document::get_doc_rels() const {
    return this->_doc_rels;
}
pugi::xml_document& Document::get_rels() const {
    return this->_rels;
}
void Document::open() {
    void *buf = nullptr;
    size_t buf_size = 0;

    // Open file and load "xml" content to the document variable
    zip_t *zip = zip_open(this->doc_file.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');

    zip_entry_open(zip, "word/document.xml");
    zip_entry_read(zip, &buf, &buf_size);
    zip_entry_close(zip);
    this->_document.load_buffer(buf, buf_size);
    free(buf);

    buf = nullptr;
    buf_size = 0;
    zip_entry_open(zip, "word/_rels/document.xml.rels");
    zip_entry_read(zip, &buf, &buf_size);
    zip_entry_close(zip);
    this->_doc_rels.load_buffer(buf, buf_size);
    free(buf);

    buf = nullptr;
    buf_size = 0;
    zip_entry_open(zip, "_rels/.rels");
    zip_entry_read(zip, &buf, &buf_size);
    zip_entry_close(zip);
    this->_rels.load_buffer(buf, buf_size);
    free(buf);

    zip_close(zip);

    this->_body = this->_document.child("w:document").child("w:body");
}
bool Document::save() const {
    // minizip only supports appending or writing to new files
    // So we must:
    // - make a new file
    // - write any new files
    // - copy the old files
    // - delete old docx
    // - rename new file to old file

    // Read document buffer
    xml_string_writer doc_writer, doc_rels_writer, rels_writer;
    this->_document.print(doc_writer);
    //std::cout << doc_writer.result.c_str() << std::endl;
    this->_doc_rels.print(doc_rels_writer);
    //std::cout << doc_rels_writer.result.c_str() << std::endl;
    this->_rels.print(rels_writer);
    //std::cout << rels_writer.result.c_str() << std::endl;

    // Open file and replace "xml" content
    std::string original_file = this->doc_file;
    std::string temp_file = this->doc_file + ".XXXXXX";
#ifdef _MSC_VER
    int sizeInChars = strlen(temp_file.c_str()) + 1;
    int err = _mktemp_s((char*)(temp_file.c_str()), sizeInChars);
#else
    temp_file = _mktemp((char*)(temp_file.c_str()));
#endif
    // Create the new file
    zip_t* new_zip = zip_open(temp_file.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');

    // Write out document.xml
    zip_entry_open(new_zip, "word/document.xml");
    const char* buf = doc_writer.result.c_str();
    zip_entry_write(new_zip, buf, strlen(buf));
    zip_entry_close(new_zip);

    // Open the original zip and copy all files which are not replaced by us
    zip_t* orig_zip = zip_open(original_file.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');

    // Loop & copy each relevant entry in the original zip
    ssize_t orig_zip_entry_ct = zip_entries_total(orig_zip);
    for (ssize_t i = 0; i < orig_zip_entry_ct; i++) {
        zip_entry_openbyindex(orig_zip, i);
        const char* name = zip_entry_name(orig_zip);

        // Skip copying the original file
        if (std::string(name) != std::string("word/document.xml")) {
            if (std::string(name) == std::string("word/_rels/document.xml.rels")) {
                zip_entry_open(new_zip, name);
                buf = doc_rels_writer.result.c_str();
                zip_entry_write(new_zip, buf, strlen(buf));
                zip_entry_close(new_zip);
            }
            else if (std::string(name) == std::string("_rels/.rels")) {
                zip_entry_open(new_zip, name);
                buf = rels_writer.result.c_str();
                zip_entry_write(new_zip, buf, strlen(buf));
                zip_entry_close(new_zip);
            }
            else {
                // Read the old content
                void* entry_buf = nullptr;
                size_t entry_buf_size = 0;
                zip_entry_read(orig_zip, &entry_buf, &entry_buf_size);
                // Write into new zip
                zip_entry_open(new_zip, name);
                zip_entry_write(new_zip, entry_buf, entry_buf_size);
                zip_entry_close(new_zip);
                free(entry_buf);
            }
        }

        zip_entry_close(orig_zip);
    }

    // `word/media/image_XXXXXX.png`
    if (!this->_medias.empty()) {
        for (auto& media : this->_medias) {
            zip_entry_open(new_zip, std::string("word/" + media.media_name).c_str());
            zip_entry_fwrite(new_zip, media.file_path.c_str());
            zip_entry_close(new_zip);
        }
    }

    // Close both zips
    zip_close(orig_zip);
    zip_close(new_zip);

    // Remove original zip, rename new to correct name
    remove(original_file.c_str());
    return rename(temp_file.c_str(), original_file.c_str());
}
std::vector<Table>& Document::tables() {
    if (_tables.empty()) {
        for(auto & it : this->_body.children()) {
            if (std::string(it.name()) == "w:tbl") {
                _tables.emplace_back(this->_body, it);
            }
        }
    }
    return _tables;
}
std::vector<Paragraph>& Document::paragraphs() {
    if (_paragraphs.empty()) {
        for(auto & it : this->_body.children()) {
            if (std::string(it.name()) == "w:p") {
                _paragraphs.emplace_back(this->_body, it);
            }
        }
    }
    return _paragraphs;
}
std::vector<MediaObject>& Document::medias() {
    return this->_medias;
}

// DocElement base class
DocElement::DocElement(const pugi::xml_node& parent, const pugi::xml_node& current) {
    this->parent = parent;
    this->current = current;
}
void DocElement::set_parent(const pugi::xml_node& node) {
    this->parent = node;
    this->current = this->parent.child("w:p");
}
pugi::xml_node& DocElement::get_parent() const {
    return this->parent;
}
void DocElement::set_current(const pugi::xml_node& node) {
    this->current = node;
}
pugi::xml_node& DocElement::get_current() const {
    return this->current;
}
DocElement& DocElement::next() {
    this->current = this->current.next_sibling();
    return *this;
}
DocElement& DocElement::prev() {
    this->current = this->current.previous_sibling();
    return *this;
}
bool DocElement::has_next() const {
    return this->current != nullptr;
}
std::vector<DocElement> DocElement::children() {
    std::vector<DocElement> _children;
    for (auto & it : this->get_current().children()) {
        _children.emplace_back(DocElement(this->get_current(), it));
    }
    return _children;
}

// Run class
void Run::set_parent(const pugi::xml_node& node) {
    this->set_current(this->get_parent().child("w:r"));
    //this->child.set_parent(this->current);
}
std::vector<Text>& Run::texts() {
    if (_texts.empty()) {
        for (auto & it : this->get_current().children()) {
            if (std::string(it.name()) == "w:t") {
                _texts.emplace_back(this->get_current(), it);
            }
        }
    }
    return _texts;
}
std::vector<Picture>& Run::pictures() {
    if (_pictures.empty()) {
        for (auto & it : this->get_current().children()) {
            if (std::string(it.name()) == "w:drawing") {
                _pictures.emplace_back(this->get_current(), it);
            }
        }
    }
    return _pictures;
}
bool Run::add_text(const std::string& text) {
    return this->add_text(text.c_str());
}
bool Run::add_text(const char* text) {
#ifdef _WIN32
    std::string utf_str;
    GB2312ToUTF8(text, utf_str);
    for (auto& it : this->get_current().children()) {
        if (std::string(it.name()) == "w:t") {
            return it.text().set(utf_str.c_str());
        }
    }
    return this->get_current().append_child("w:t").text().set(utf_str.c_str());
#else
    for (auto& it : this->get_current().children()) {
        if (std::string(it.name()) == "w:t") {
            return it.text().set(text);
        }
}    }
    return this->get_current().append_child("w:t").text().set(text);
#endif
}
bool Run::add_picture(Document& doc, const std::string& srcImg, int wd, int ht) {
    return this->add_picture(doc, srcImg.c_str(), wd, ht);
}
bool Run::add_picture(Document& doc, const char* srcImg, int wd, int ht) {
    //`word/document.xml`, cx="{0}" cy="{1}" embed="{2}" name="{3}" descr="{4}"
    // ooxml uses image size in EMU : 
          // image in inches(in) is : pt / 72
          // image in EMU is : in * 914400
    int cx = (int) (wd * 12700);
    int cy = (int) (ht * 12700);
    int i = 1;
    for (auto & doc_rel : doc.get_doc_rels().child("Relationships").children("Relationship")) {
        //std::string rid = doc_rel.attribute("Id").value();
        i++;
    }
    std::string embed = "rId" + std::to_string(i); //"rId" + uid

    i = 1;
    //std::random_device rd; //BUG: there is memory leak!!!
    //unsigned int r = rd();
    unsigned int r = dice();
    std::string pic_name = "media/image_" + std::to_string(r) + ".png";
    for (pugi::xml_node doc_rel : doc.get_doc_rels().child("Relationships").children("Relationship")) {
        std::string imageType = doc_rel.attribute("Type").value();
        if (imageType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
            //unsigned int rn = rd();
            unsigned int rn = dice();
            pic_name = "media/image_" + std::to_string(rn) + ".png";
            std::string target = doc_rel.attribute("Target").value();
            while (target == pic_name) {
                //unsigned int rw = rd();
                unsigned int rw = dice();
                pic_name = "media/image_" + std::to_string(rw) + ".png";
            }
            i++;
        }
    }
    std::string name = "Picture " + std::to_string(i); //"Picture " + uid
    std::string descr = "";
    auto this_run = this->get_current();
    auto drawing = this_run.append_child("w:drawing");
    auto wpinline = drawing.append_child("wp:inline");
    wpinline.append_attribute("distT").set_value("0");
    wpinline.append_attribute("distB").set_value("0");
    wpinline.append_attribute("distL").set_value("0");
    wpinline.append_attribute("distR").set_value("0");
    wpinline.append_attribute("xmlns:wp").set_value("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
    auto wpextent = wpinline.append_child("wp:extent");
    wpextent.append_attribute("cx").set_value(std::to_string(cx).c_str());
    wpextent.append_attribute("cy").set_value(std::to_string(cy).c_str());
    auto wpeffectExtent = wpinline.append_child("wp:effectExtent");
    wpeffectExtent.append_attribute("l").set_value("0");
    wpeffectExtent.append_attribute("t").set_value("0");
    wpeffectExtent.append_attribute("r").set_value("0");
    wpeffectExtent.append_attribute("b").set_value("0");
    auto wpdocPr = wpinline.append_child("wp:docPr");
    wpdocPr.append_attribute("id").set_value("0");
    wpdocPr.append_attribute("name").set_value(name.c_str());
    wpdocPr.append_attribute("descr").set_value(descr.c_str());
    auto wpcNvGraphicFramePr = wpinline.append_child("wp:cNvGraphicFramePr");
    auto agraphicFrameLocks = wpcNvGraphicFramePr.append_child("a:graphicFrameLocks");
    agraphicFrameLocks.append_attribute("xmlns:a").set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
    agraphicFrameLocks.append_attribute("noChangeAspect").set_value("1");
    auto agraphic = wpinline.append_child("a:graphic");
    agraphic.append_attribute("xmlns:a").set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
    auto agraphicData = agraphic.append_child("a:graphicData");
    agraphicData.append_attribute("uri").set_value("http://schemas.openxmlformats.org/drawingml/2006/picture");
    auto picpic = agraphicData.append_child("pic:pic");
    picpic.append_attribute("xmlns:pic").set_value("http://schemas.openxmlformats.org/drawingml/2006/picture");
    auto picnvPicPr = picpic.append_child("pic:nvPicPr");
    auto piccNvPr = picnvPicPr.append_child("pic:cNvPr");
    piccNvPr.append_attribute("id").set_value("0");
    piccNvPr.append_attribute("name").set_value(name.c_str());
    auto piccNvPicPr = picnvPicPr.append_child("pic:cNvPicPr");
    auto apicLocks = piccNvPicPr.append_child("a:picLocks");
    apicLocks.append_attribute("noChangeAspect").set_value("1");
    auto picblipFill = picpic.append_child("pic:blipFill");
    auto ablip = picblipFill.append_child("a:blip");
    ablip.append_attribute("r:embed").set_value(embed.c_str());
    ablip.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    auto astretch = picblipFill.append_child("a:stretch");
    auto afillRect = astretch.append_child("a:fillRect");
    auto picspPr = picpic.append_child("pic:spPr");
    auto axfrm = picspPr.append_child("a:xfrm");
    auto aoff = axfrm.append_child("a:off");
    aoff.append_attribute("x").set_value("0");
    aoff.append_attribute("y").set_value("0");
    auto aext = axfrm.append_child("a:ext");
    aext.append_attribute("cx").set_value(std::to_string(cx).c_str());
    aext.append_attribute("cy").set_value(std::to_string(cy).c_str());
    auto aprstGeom = picspPr.append_child("a:prstGeom");
    aprstGeom.append_attribute("prst").set_value("rect");
    auto aavLst = aprstGeom.append_child("a:avLst");

    //`word/_rels/document.xml.rels`
    pugi::xml_node doc_rid = doc.get_doc_rels().child("Relationships").append_child("Relationship");
    doc_rid.append_attribute("Id").set_value(embed.c_str());
    doc_rid.append_attribute("Type").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    doc_rid.append_attribute("Target").set_value(pic_name.c_str());

    //`_rels/.rels`
    pugi::xml_node rid = doc.get_rels().child("Relationships").append_child("Relationship");
    rid.append_attribute("Id").set_value(embed.c_str());
    rid.append_attribute("Type").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    rid.append_attribute("Target").set_value("word/document.xml");

    //`word\{target}` target=media\imageXX.png
    // append `document.MediaObject`
    MediaObject mo(std::string(srcImg), pic_name);
    doc.medias().emplace_back(mo);

    return true;
}

// Text class
void Text::set_parent(const pugi::xml_node& node) {
    this->set_current(this->get_parent().child("w:t"));
    //this->child.set_parent(this->current);
}
std::string Text::get_text() const {
    return this->get_current().text().get();
}
bool Text::set_text(const std::string& text) const {
    return this->set_text(text.c_str());
}
bool Text::set_text(const char* text) const {
#ifdef _WIN32
    std::string utf_str;
    GB2312ToUTF8(text, utf_str);
    return this->get_current().text().set(utf_str.c_str());
#else
    return this->get_current().text().set(text);
#endif // _WIN32
}

// Picture class
void Picture::set_parent(const pugi::xml_node& node) {
    this->set_current(this->get_parent().child("w:drawing"));
    //this->child.set_parent(this->current);
}

// TableCell class
void TableCell::set_parent(const pugi::xml_node& node) {
    this->set_current(this->get_parent().child("w:tc"));
    //this->child.set_parent(this->current);
}
std::vector<Paragraph>& TableCell::paragraphs() {
    if (_paragraphs.empty()) {
        for (auto & it : this->get_current().children()) {
            if (std::string(it.name()) == "w:p") {
                _paragraphs.emplace_back(this->get_current(), it);
            }
        }
    }
    return _paragraphs;
}

// TableRow class
void TableRow::set_parent(const pugi::xml_node& node) {
    this->set_current(this->get_parent().child("w:tr"));
    //this->child.set_parent(this->current);
}
std::vector<TableCell>& TableRow::cells() {
    if (_cells.empty()) {
        for (auto & it : this->get_current().children()) {
            if (std::string(it.name()) == "w:tc") {
                _cells.emplace_back(this->get_current(), it);
            }
        }
    }
    return _cells;
}
TableRow TableRow::append_row(bool copyContent) {
    pugi::xml_node new_row = this->get_parent().append_copy(this->get_current());
    if (!copyContent) {
        for (auto& tc : new_row.children()) {
            if (std::string(tc.name()) == "w:tc") {
                for (auto& p : tc.children()) {
                    if (std::string(p.name()) == "w:p") {
                        for (auto& r : p.children()) {
                            if (std::string(r.name()) == "w:r") {
                                r.remove_children();
                            }
                        }
                    }
                }
            }
        }
    }
    return TableRow{this->get_parent(), new_row};
}

// Table class
void Table::set_parent(const pugi::xml_node& node) {
    this->set_current(this->get_parent().child("w:tbl"));
    //this->child.set_parent(this->current);
}
std::vector<TableRow>& Table::rows() {
    if (_rows.empty()) {
        for (auto & it : this->get_current().children()) {
            if (std::string(it.name()) == "w:tr") {
                _rows.emplace_back(this->get_current(), it);
            }
        }
    }
    return _rows;
}

// Paragraph class
void Paragraph::set_parent(const pugi::xml_node& node) {
    this->set_current(this->get_parent().child("w:p"));
    //this->child.set_parent(this->current);
}
std::vector<Run>& Paragraph::runs() {
    if (_runs.empty()) {
        for (auto & it : this->get_current().children()) {
            if (std::string(it.name()) == "w:r") {
                _runs.emplace_back(this->get_current(), it);
            }
        }
    }
    if (_runs.empty()) {
        auto r = this->get_current().append_child("w:r");
        _runs.emplace_back(this->get_current(), r);
    }
    return _runs;
}
