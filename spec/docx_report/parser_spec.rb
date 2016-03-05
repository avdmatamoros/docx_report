require 'spec_helper'

describe DocxReport::Parser do
  before do
    @doc = DocxReport::Document.new('spec/files/template.docx')
  end

  subject { DocxReport::Parser.new @doc }

  it 'replaces text fields' do
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "{@name}") ]').count).to eq(1)
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "Ahmed") ]').count).to eq(0)
    subject.replace_all_fields({ 'name' => 'Ahmed' })
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "{@name}") ]').count).to eq(0)
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "Ahmed") ]').count).to eq(1)
  end

  it 'fill table fields' do
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "{@title}") ]').count).to eq(1)
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "First record") ]').count).to eq(0)
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "Third record") ]').count).to eq(0)
    table = DocxReport::Table.new 'table1'
    table.new_record.add_field 'title', 'First record'
    table.new_record.add_field 'title', 'Second record'
    table.new_record.add_field 'title', 'Third record'
    subject.fill_all_tables([table])
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "{@title}") ]').count).to eq(0)
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "First record") ]').count).to eq(1)
    expect(@doc.files['word/document.xml']
               .xpath('//*[contains(text(), "Third record") ]').count).to eq(1)
  end
end
