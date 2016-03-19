require 'spec_helper'

describe DocxReport::Record do
  subject { DocxReport::Record.new }

  it 'adds text fields' do
    subject.add_field 'name', 'Ahmed Abudaqqa'
    expect(subject.fields.detect do |f|
      f.name == '@name@' && f.value == 'Ahmed Abudaqqa'
    end).to_not be_nil
  end
end
