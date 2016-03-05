require 'spec_helper'

describe DocxReport::Report do
  subject { DocxReport::Report.new 'spec/files/template.docx' }

  it 'adds text fields' do
    subject.add_field 'name', 'Ahmed Abudaqqa'
    expect(subject.fields).to eq({'{@name}' => 'Ahmed Abudaqqa'})
  end

  it 'adds table loaded form a collection of data' do
    items = [{ name: 'Item 1', details: 'details of item 1' },
             { name: 'Item 2', details: 'details of item 2' },
             { name: 'Item 3', details: 'details of item 3' }]
    subject.add_table 'name', items do |table|
      table.add_field(:title, :name)
      table.add_field(:description, :details)
    end

    records = subject.tables.first.records
    expect(records.count).to eq(3)
    expect(records.first.fields).to eq(
      {
        '{@title}' => 'Item 1',
        '{@description}' => 'details of item 1'
      })
    expect(records.last.fields).to eq(
      {
        '{@title}' => 'Item 3',
        '{@description}' => 'details of item 3'
      })
  end

  it 'generates new docx file' do
    subject.generate_docx 'output.docx'
    expect(File.exists? 'output.docx').to be true
    File.delete 'output.docx'
  end
end
