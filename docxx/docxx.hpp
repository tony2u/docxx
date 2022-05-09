#ifndef DOCXX_LIBRARY_H
#define DOCXX_LIBRARY_H

#include <string>
#include <utility>
#include <vector>
#include "pugixml.hpp"
#include "zip.h"

namespace docxx
{
    struct MediaObject
    {
        std::string file_path;  // full path include file name
        std::string media_name; // media name saved into `word\media` directory
        MediaObject(std::string&& fp, std::string&& name) : file_path(std::move(fp)), media_name(std::move(name)) { }
        MediaObject(const std::string& fp, const std::string& name) : file_path(fp), media_name(name) { }
    };

    // Store both the document and xml so that
    // they can be accessed by derived types
    class DocElement
    {
    private:
        //Document document;
        mutable pugi::xml_node parent;
        mutable pugi::xml_node current;
    protected:
        DocElement() = default;
    public:
        DocElement(const pugi::xml_node& parent, const pugi::xml_node& current);
        ~DocElement() = default;
        virtual void set_parent(const pugi::xml_node&);
        virtual pugi::xml_node& get_parent() const;
        virtual void set_current(const pugi::xml_node&);
        virtual pugi::xml_node& get_current() const;
        virtual DocElement& next();
        virtual DocElement& prev();
        virtual bool has_next() const;
        //all children elements
        virtual std::vector<DocElement> children();
    };

    // Represents a document Text object
    class Text : public DocElement
    {
    private:
    protected:
    public:
        Text(const pugi::xml_node& parent, const pugi::xml_node& current) : DocElement(parent, current) {}
        ~Text() = default;
        void set_parent(const pugi::xml_node&) override;
        std::string get_text() const;
        bool set_text(const std::string&) const;
        bool set_text(const char*) const;
    };

    // Represents a document Picture object
    class Picture : public DocElement
    {
    private:
    protected:
    public:
        Picture(const pugi::xml_node& parent, const pugi::xml_node& current) : DocElement(parent, current)/*, _flag(0)*/ {}
        ~Picture() = default;
        void set_parent(const pugi::xml_node&) override;
    };

    // Represents a document Run object
    class Document;
    class Run : public DocElement
    {
    private:
        std::vector<Text> _texts;
        std::vector<Picture> _pictures;
    protected:
    public:
        Run(const pugi::xml_node &parent, const pugi::xml_node &current) : DocElement(parent, current) {}
        ~Run() = default;
        void set_parent(const pugi::xml_node&) override;
        bool add_text(const std::string&, bool);
        bool add_text(const char*, bool);
        bool add_picture(Document&, const std::string&, int, int);
        bool add_picture(Document&, const char*, int, int);
        std::vector<Text>& texts();
        std::vector<Picture>& pictures();
    };

    // Represents a document paragraph
    class Paragraph : public DocElement
    {
    private:
        std::vector<Run> _runs;
    protected:
    public:
        Paragraph(const pugi::xml_node &parent, const pugi::xml_node &current) : DocElement(parent, current) {}
        ~Paragraph() = default;
        void set_parent(const pugi::xml_node&) override;
        std::vector<Run>& runs();
    };

    // Represents a document TableCell object
    // TableCell contains one or more paragraphs
    class TableCell : public DocElement
    {
    private:
        std::vector<Paragraph> _paragraphs;
    protected:
    public:
        TableCell(const pugi::xml_node& parent, const pugi::xml_node& current) : DocElement(parent, current) {}
        ~TableCell() = default;
        void set_parent(const pugi::xml_node&) override;
        std::vector<Paragraph>& paragraphs();
    };

    // Represents a document TableRow object
    // TableRow consists of one or more TableCells
    class TableRow : public DocElement
    {
    private:
        std::vector<TableCell> _cells;
    protected:
    public:
        TableRow(const pugi::xml_node& parent, const pugi::xml_node& current) : DocElement(parent, current) {}
        ~TableRow() = default;
        void set_parent(const pugi::xml_node&) override;
        std::vector<TableCell>& cells();
        TableRow append_row(bool copyContent);
    };

    // Represents a document Table object
    // Table consists of one or more TableRows
    class Table : public DocElement
    {
    private:
        std::vector<TableRow> _rows;
    protected:
    public:
        Table(const pugi::xml_node& parent, const pugi::xml_node& current) : DocElement(parent, current) {}
        ~Table() = default;
        void set_parent(const pugi::xml_node&) override;
        std::vector<TableRow>& rows();
    };

    // Document contains whole the docx file
    // and stores paragraphs
    class Document
    {
    private:
        std::string doc_file;
        mutable pugi::xml_node _body;
        mutable pugi::xml_document _document; //`word/document.xml`
        mutable pugi::xml_document _doc_rels; //`word/_rels/document.xml.rels`
        mutable pugi::xml_document _rels;     //`_rels/.rels`
        std::vector<Table> _tables;
        std::vector<Paragraph> _paragraphs;
        std::vector<MediaObject> _medias;
        //friend bool Run::add_picture(Document&, const char*, int, int);
    protected:
    public:
        Document() = default;
        Document(const std::string&&, bool);
        Document(const char*, bool);
        ~Document() = default;
        //Document(Document&);
        pugi::xml_node& get_body() const;
        pugi::xml_document& get_document() const;
        pugi::xml_document& get_doc_rels() const;
        pugi::xml_document& get_rels() const;
        void open();
        bool save() const;

        std::vector<Table>& tables();
        std::vector<Paragraph>& paragraphs();
        std::vector<MediaObject>& medias();
    };

}

#endif // DOCXX_LIBRARY_H
