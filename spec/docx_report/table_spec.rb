require 'spec_helper'

describe DocxReport::Table do
  subject { DocxReport::Table.new 'test' }

  it 'adds records' do
    subject.new_record.add_field 'name', 'Ahmed Abudaqqa'
    expect(subject.records.first.fields.detect do |f|
      f.name == '@name@' && f.value == 'Ahmed Abudaqqa'
    end).to_not be_nil
  end
end
