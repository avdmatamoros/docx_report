require 'spec_helper'

describe DocxReport::Table do
  subject { DocxReport::Table.new 'test' }

  it 'adds records' do
    subject.new_record.add_field 'name', 'Ahmed Abudaqqa'
    expect(subject.records.first.fields).to eq('{@name}' => 'Ahmed Abudaqqa')
  end
end
