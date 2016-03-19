Gem::Specification.new do |s|
  s.name        = 'docx_report'
  s.version     = '0.1.0'
  s.date        = '2016-03-20'
  s.summary     = 'docx_report generate docx files based on previously created
                   .docx template file'
  s.description = 'docx_report is a light weight gem that generates docx files
                   by replacing strings and inserting images on previously
                   created .docx template file'
  s.authors     = ['Ahmed Abudaqqa']
  s.email       = 'ahmed@abudaqqa.com'
  s.files       = `git ls-files`.split("\n").select { |f| f.match(%r{^(lib)/}) }
  s.homepage    = 'https://github.com/abudaqqa/docx_report'
  s.license     = 'MIT'

  s.add_dependency 'nokogiri', '~> 1.6'
  s.add_dependency 'rubyzip',  '~> 1.2'
end
