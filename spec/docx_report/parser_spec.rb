require 'spec_helper'

describe DocxReport::Parser do
  before do
    @doc = DocxReport::Document.new('spec/files/template.docx')
  end

  subject { DocxReport::Parser.new @doc }

  it 'replaces text fields' do
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@name@") ]').count).to eq(1)
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Ahmed") ]').count).to eq(0)
    subject.replace_all_fields([DocxReport::Field.new('name', 'Ahmed')])
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@name@") ]').count).to eq(0)
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Ahmed") ]').count).to eq(1)
  end

  it 'replaces hyperlinks fields' do
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="@url@"]').count).to eq(1)
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="abudaqqa.com"]').count).to eq(0)
    subject.replace_all_fields([DocxReport::Field.new('url', 'abudaqqa.com',
                                                      :hyperlink)])
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="@url@"]').count).to eq(0)
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="abudaqqa.com"]').count).to eq(1)
  end

  it 'fill table fields' do
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@title@") ]').count).to eq(1)
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "First record") ]').count).to eq(0)
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Third record") ]').count).to eq(0)
    table = DocxReport::Table.new 'table1'
    table.new_record.add_field 'title', 'First record'
    table.new_record.add_field 'title', 'Second record'
    table.new_record.add_field 'title', 'Third record'
    subject.fill_all_tables [table]
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@title@") ]').count).to eq(0)
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "First record") ]').count).to eq(1)
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Third record") ]').count).to eq(1)
  end
end
