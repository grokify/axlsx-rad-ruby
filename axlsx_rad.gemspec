Gem::Specification.new do |s|
  s.name        = 'axlsx_rad'
  s.version     = '0.0.2'
  s.date        = '2014-03-23'
  s.summary     = 'AxlsxRad - Rapid XLSX Writer for Customized Spreadsheets'
  s.description = 'Efficiently create styled spreadsheets using JsonDoc::Document objects'
  s.authors     = ['John Wang']
  s.email       = 'john@johnwang.com'
  s.files       = [
    'CHANGELOG.md',
    'LICENSE',
    'README.md',
    'Rakefile',
    'VERSION',
    'lib/axlsx_rad.rb',
    'lib/axlsx_rad/config.rb',
    'lib/axlsx_rad/workbook.rb',
    'test/test_setup.rb'
  ]
  s.homepage    = 'http://johnwang.com/'
  s.license     = 'MIT'
  s.add_runtime_dependency 'axlsx', '~> 1.3', '>= 1.3.5'
end
