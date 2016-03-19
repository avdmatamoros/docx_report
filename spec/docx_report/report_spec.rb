require 'spec_helper'

describe DocxReport::Report do
  subject { DocxReport::Report.new 'spec/files/template.docx' }

  it 'adds text fields' do
    subject.add_field 'name', 'Ahmed Abudaqqa'
    expect(subject.fields.detect do |f|
      f.name == '@name@' && f.value == 'Ahmed Abudaqqa' && f.type == :text
    end).to_not be_nil
  end

  it 'adds hyperlink fields' do
    subject.add_field 'url', 'abudaqqa.com', :hyperlink
    expect(subject.fields.detect do |f|
      f.name == '@url@' && f.value == 'abudaqqa.com' && f.type == :hyperlink
    end).to_not be_nil
  end

  it 'adds table loaded form a collection of data' do
    items = [{ name: 'Item 1', details: 'details of item 1', url: 'web1.com' },
             { name: 'Item 2', details: 'details of item 2', url: 'web2.com' },
             { name: 'Item 3', details: 'details of item 3', url: 'web3.com' }]
    subject.add_table 'table1', items do |table|
      table.add_field(:title, :name)
      table.add_field(:description) { |item| "Details: #{item[:details]}" }
      table.add_field(:url, :url, :hyperlink)
    end
    records = subject.tables.first.records
    expect(records.count).to eq(3)
    expect(records.detect do |r|
      r.fields.any? { |f| f.name == '@title@' && f.value == 'Item 1' } &&
      r.fields.any? do |f|
        f.name == '@description@' && f.value == 'Details: details of item 1'
      end &&
      r.fields.any? do |f|
        f.name == '@url@' && f.value == 'web1.com' && f.type == :hyperlink
      end
    end).not_to be_nil
    expect(records.detect do |r|
      r.fields.any? { |f| f.name == '@title@' && f.value == 'Item 3' } &&
      r.fields.any? do |f|
        f.name == '@description@' && f.value == 'Details: details of item 3'
      end &&
      r.fields.any? do |f|
        f.name == '@url@' && f.value == 'web3.com' && f.type == :hyperlink
      end
    end).not_to be_nil
  end

  it 'generates new docx file after apply changes' do
    temp = Tempfile.new 'output.docx'
    subject.add_field 'photo', 'spec/files/ruby.png', :image
    items = [{ name: 'Item 1', details: 'details of item 1', url: 'web1.com',
               link: 'https://www.ruby-lang.org/images/header-ruby-logo@2x.png',
               subitems: [{ subtitle: 'Item4' }, { subtitle: 'Item5' }] },
             { name: 'Item 2', details: 'details of item 2', url: 'web2.com',
               link: 'spec/files/ruby.png',
               subitems: [{ subtitle: 'Item6' }] },
             { name: 'Item 3', details: 'details of item 3', url: 'web3.com',
               link: 'spec/files/ruby.png',
               subitems: [{ subtitle: 'Item7' }, { subtitle: 'Item8' },
                          { subtitle: 'Item9' }] }]
    subject.add_table 'table1', items do |table|
      table.add_field(:title, :name)
      table.add_field(:description) { |item| "Details: #{item[:details]}" }
      table.add_field(:url, :url, :hyperlink)
      table.add_field(:picture, :link, :image)
      table.add_table(:subtable, :subitems).add_field(:subtitle, :subtitle)
    end
    subject.generate_docx temp.path
    expect(File.exist?(temp.path)).to be true
    temp.close!
  end

  private

  def load_main_xml(filename)
    zip = Zip::File.open(filename)
    main = Nokogiri::XML zip.read('word/document.xml')
    zip.close
    main
  end
end
