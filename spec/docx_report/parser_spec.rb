require 'spec_helper'

describe DocxReport::Parser do
  before do
    @doc = DocxReport::Document.new('spec/files/template.docx')
  end

  subject { DocxReport::Parser.new @doc }

  it 'replaces text fields' do
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@name@") ]').any?).to be true
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Ahmed") ]').any?).to be false
    subject.replace_all_fields([DocxReport::Field.new('name', 'Ahmed')])
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@name@") ]').any?).to be false
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Ahmed") ]').any?).to be true
  end

  it 'replaces hyperlinks fields' do
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="@url@"]').any?).to be true
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="abudaqqa.com"]').any?).to be false
    subject.replace_all_fields([DocxReport::Field.new('url', 'abudaqqa.com',
                                                      :hyperlink)])
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="@url@"]').any?).to be false
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="abudaqqa.com"]').any?).to be true
  end

  it 'replaces images fields' do
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="media/image2.%s"]').any?)
           .to be false
    subject.replace_all_fields([DocxReport::Field.new('photo',
                                                      'spec/files/ruby.png',
                                                      :image)])
    expect(@doc.files.detect do |f|
             f.rels_name == 'word/_rels/document.xml.rels'
           end.rels_xml.xpath('//*[@Target="media/image2.%s"]').any?)
           .to be true
  end

  it 'fill table fields' do
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@title@") ]').any?).to be true
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "First record") ]').any?)
               .to be false
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Third record") ]').any?)
               .to be false
    table = DocxReport::Table.new 'table1'
    table.new_record.add_field 'title', 'First record'
    table.new_record.add_field 'title', 'Second record'
    table.new_record.add_field 'title', 'Third record'
    subject.fill_all_tables [table]
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "@title@") ]').any?).to be false
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "First record") ]').any?).to be true
    expect(@doc.files.detect { |f| f.name == 'word/document.xml' }.xml
               .xpath('//*[contains(text(), "Third record") ]').any?).to be true
  end
end
