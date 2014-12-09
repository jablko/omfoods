#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import division, unicode_literals

import cgi, csv, htmlentitydefs, os, re, shutil, sys, time
from HTMLParser import HTMLParser
from xml.etree import ElementTree as etree

sys.path.insert(0, 'XlsxWriter')
import xlsxwriter

if os.environ['REQUEST_METHOD'] == 'POST':
  filename = time.strftime('%Y-%m-%dT%H:%M:%S')

  form = cgi.FieldStorage()
  shutil.copyfileobj(form['file'].file, open(filename + '.csv', 'w'))

  class parser(HTMLParser):
    def handle_charref(self, name):
      if name[0] in 'Xx':
        self.data += unichr(int(name[1:], 16))

      else:
        self.data += unichr(int(name))

    def handle_data(self, data):
      self.data += data

    def handle_entityref(self, name):
      if name in htmlentitydefs.name2codepoint:
        self.data += unichr(htmlentitydefs.name2codepoint[name])

      else:
        self.data += '&' + name + ';'

  parser = parser()

  country_codes = {
    'cote d\'ivore': 'CI',
    'itali': 'IT',
    'paraguy': 'PY',
    'rajathan': 'IN-RJ',
    'sicily': 'IT-82',
    'south korea': 'KR',
    'tanzania': 'TZ',
    'us': 'US',
    'philippine': 'PH',
    }

  tree = etree.parse('/usr/share/xml/iso-codes/iso_3166.xml')
  for entry in tree.findall('iso_3166_entry'):

    code = entry.get('alpha_2_code')
    name = entry.get('common_name') or entry.get('name')

    country_codes[name.lower()] = code
    country_codes[name.replace(' ', '').lower()] = code

  countries = list(country_codes)
  countries.sort(key=len)
  countries.reverse()
  countries = '|'.join(map(re.escape, countries))
  origin_re = re.compile('''
    (?:
      Origin:\ *
      |
      (?:Made\ in|Product\ of)(?::\ *|\ (?:the\ )?)
    )
    (''' + countries + ''')
    (?:
      (?:\ or\ |\ *[&,/]\ *)
      (''' + countries + ''')
    )?
    ''', re.IGNORECASE|re.VERBOSE)

  elide_re = re.compile('[^0-9a-z]+', re.IGNORECASE)
  def fixup_groups(groups):
    head, value, units, tail = groups

    if head:

      head = elide_re.sub(' ', head)
      head = head.lower()

    if value:

      value = value.replace(',', '')
      value = reduce(lambda a, b: a / b, map(float, value.split('/')))
      value = {
        1/2: '1/2',
        1/3: '1/3',
        2/3: '2/3',
        1/4: '1/4',
        3/4: '3/4',
        1/8: '1/8',
        }.get(value) or format(value, ',.2f').rstrip('0').rstrip('.')

    if units:

      units = units.lower()
      if units == 'gallon':
        units = 'gal'

      elif units == 'l':
        units = 'L'

      elif units == 'lbs':
        units = 'lb'

      elif units not in ['ct', 'g', 'gal', 'kg', 'lb', 'ml', 'oz']:
        units = ' ' + units

    if tail == 'x':
      tail = ''

    elif tail:

      tail = elide_re.sub(' ', tail)
      tail = ' ' + tail.lower()

    return head + value + units + tail

  not_numbers_or_letters = '(?:[^.0-9a-z]|\.(?![0-9]))*'
  not_numbers_LETTERS = '[^0-9]*[a-z]'
  not_numbers = '(?:[^.0-9]|\.(?![0-9]))*'

  # The expression that precedes this should swallow any initial commas
  number = '''
    (?:,?[0-9]+)*\.[0-9]+ # Non-integer
    |
    (?:,?[0-9]+)+(?:\ */\ *[0-9]+(?:,[0-9]+)*)? # Integer or fraction
    '''

  spaces_LETTERS = '(?:\ *[a-z])*'
  LETTERS_spaces = '(?:[a-z]\ *)*'
  head_re = re.compile('''
    ''' + not_numbers_or_letters + '''
    (''' + not_numbers_LETTERS + ''')? # head
    ''' + not_numbers + '''
    (?:
      (''' + number + ''') # value
      \ *(?:th\ +)?
      ([a-z]*) # units
      \ *
      (''' + spaces_LETTERS + ''') # tail
    )?
    ''', re.IGNORECASE|re.VERBOSE)
  tail_re = re.compile(LETTERS_spaces + '(?:[0-9](?! *(?:ct|g|gal|gallon|kg|l|lb|lbs|ml|oz)(?![a-z]))[ a-z]*)+$', re.IGNORECASE)
  size_re = re.compile('''
    ''' + not_numbers_or_letters + '''
    (''' + not_numbers_LETTERS + ''')? # head
    ''' + not_numbers + '''
    (?:
      (''' + number + ''') # value
      \ *(?:th\ +)?
      ([a-z]*) # units
      ''' + not_numbers_or_letters + '''
      (''' + not_numbers_LETTERS + '''(?=\ *(?:[^ .0-9a-z]|\.(?![0-9])|$)))? # tail
    )?
    ''', re.IGNORECASE|re.VERBOSE)

  def fixup_size(size):
    if size:
      head_match = head_re.search(size, 9)

      tail_match = tail_re.search(size, head_match.end())
      if tail_match:
        parts = size_re.findall(size, head_match.end(), tail_match.start())

      else:
        parts = size_re.findall(size, head_match.end())

      head, value, units, tail = head_match.groups('')
      size = fixup_groups((head, value, units, tail))

      parts = filter(any, parts)
      if parts:
        size += ' (' + ' - '.join(map(fixup_groups, parts)) + ')'

      if tail_match:
        size = tail_match.group().lower() + ' × ' + size

      return size[0].upper() + size[1:]

  data = {}
  for row in csv.reader(open(filename + '.csv', 'U')):
    if row[0] == 'Product':
      _, name, code, description, price, allow_purchases, category = row

      name = name.decode('latin1').replace('\x89\xdb\xa2', '•')
      if name.lower().startswith('organic '):
        name = name[8:]

      code = code.decode('latin1').replace('\x89\xdb\xa2', '•')
      code = code.replace(' ', '')

      parser.data = ''
      parser.feed(description)
      description = ' '.join(parser.data.split())

      if 'certified organic' in description.lower():
        certified_organic = 'Y'

      else:
        certified_organic = None

      match = origin_re.search(description)
      if match:

        origin = country_codes[match.group(1).lower()]
        if match.group(2):
          origin += '/' + country_codes[match.group(2).lower()]

      else:
        origin = None

      if allow_purchases == 'Y':
        price = float(price)

      else:
        price = 'O/S'

      category = category.split(b';')[0]
      category = category.decode('latin1').replace('\x89\xdb\xa2', '•')

      skus = {}
      rules = []

      if category in data:
        data[category].append((name, code, certified_organic, origin, description, price, skus, rules))

      else:
        data[category] = [(name, code, certified_organic, origin, description, price, skus, rules)]

    elif row[0] == '  SKU':
      _, size, code, _, _, _, _ = row

      size = fixup_size(size)

      code = code.decode('latin1').replace('\x89\xdb\xa2', '•')
      code = code.replace(' ', '')

      skus[size] = code

    elif row[0] == '  Rule':
      _, size, code, _, price, allow_purchases, _ = row

      size = fixup_size(size)

      code = code.decode('latin1').replace('\x89\xdb\xa2', '•')
      code = code.replace(' ', '')

      if allow_purchases == 'Y':

        price = price[7:]
        price = float(price)

      else:
        price = 'O/S'

      rules.append((size, code, price))

  country_names = {}
  for entry in tree.findall('iso_3166_entry'):

    code = entry.get('alpha_2_code')
    name = entry.get('common_name') or entry.get('name')

    country_names[code] = name

  tree = etree.parse('/usr/share/xml/iso-codes/iso_3166_2.xml')
  for entry in tree.findall('.//iso_3166_2_entry'):

    code = entry.get('code')
    name = entry.get('name')

    country_names[code] = name

  country_names['KR'] = 'South Korea'
  country_names['TZ'] = 'Tanzania'

  workbook = xlsxwriter.Workbook(filename + '.xlsx')

  bold = workbook.add_format(dict(bold=True))
  border = workbook.add_format(dict(
    bottom=4,
    left=1,
    right=1))
  bottom = workbook.add_format(dict(
    bottom=1,
    left=1,
    right=1))
  center = workbook.add_format(dict(align='center'))
  header = workbook.add_format(dict(
    align='center',
    bold=True,
    bottom=1))
  header_rotation = workbook.add_format(dict(
    align='center',
    font_size=10,
    rotation=-90,
    valign='vcenter'))
  heading = workbook.add_format(dict(
    bold=True,
    font_color='#008000',
    font_size=14))
  heading_category = workbook.add_format(dict(
    align='center',
    bold=True,
    bottom=1,
    font_size=14))
  last_updated = workbook.add_format(dict(font_color='#FF0000'))
  last_updated_bold = workbook.add_format(dict(
    bold=True,
    font_color='#FF0000'))
  price = workbook.add_format(dict(num_format='$#,##0.00'))
  text_wrap = workbook.add_format(dict(text_wrap=True))
  thank_you = workbook.add_format(dict(
    align='center',
    bold=True,
    font_color='#FF0000',
    font_size=14))
  title = workbook.add_format(dict(
    align='center',
    bold=True,
    font_size=24))
  url = workbook.add_format(dict(
    align='center',
    bold=True,
    font_color='#FF0000',
    font_size=14))

  worksheet = workbook.add_worksheet()
  row = 0

  worksheet.insert_image(row, 0, 'logo.png', dict(url='http://www.omfoods.com/'))
  row += 1

  worksheet.merge_range(row, 0, row, 7, 'ORGANIC MATTERS', title)
  row += 1

  worksheet.merge_range(row, 0, row, 7, None)
  worksheet.write_url(row, 0, 'http://www.omfoods.com/', url, 'www.omfoods.com')
  row += 1

  worksheet.merge_range(row, 0, row, 7, None)
  worksheet.write_rich_string(row, 0, bold, 'CONTACT:', ' 250-505-2272 • info@omfoods.com', center)
  row += 1

  worksheet.merge_range(row, 0, row, 7, 'PO Box 1221 • Nelson BC • V1L 6H3', center)
  row += 1

  worksheet.merge_range(row, 0, row, 7, None)
  worksheet.write_rich_string(row, 0, bold, 'WAREHOUSE AND OFFICE:', ' 3505 Highway 6 • Nelson BC • V1L 6Z4', center)
  row += 1

  worksheet.merge_range(row, 0, row, 7, None)
  worksheet.write_rich_string(row, 0, bold, 'PICK-UP HOURS:', ' Monday To Friday • 10 AM TO 5:30 PM', center)
  row += 1

  worksheet.merge_range(row, 0, row, 7, None)
  worksheet.write_rich_string(row, 0, last_updated_bold, 'LAST UPDATED:', last_updated, ' ' + time.strftime('%B %-d, %Y'), center)
  row += 1

  worksheet.set_column(1, 1, 4, center)
  worksheet.set_column(2, 2, 48)
  worksheet.set_column(3, 3, 6, center)
  worksheet.set_column(4, 4, 16, center)
  worksheet.set_column(5, 5, None, price)
  worksheet.set_column(6, 6, 4)

  worksheet.set_row(row, None, header)
  worksheet.write(row, 0, 'CODE')
  worksheet.write(row, 1, 'CERTIFIED\nORGANIC', header_rotation)
  worksheet.write(row, 2, 'DESCRIPTION')
  worksheet.write(row, 3, 'ORIGIN', header_rotation)
  worksheet.write(row, 4, 'SIZE')
  worksheet.write(row, 5, 'PRICE')
  worksheet.write(row, 6, 'TAX', header_rotation)
  worksheet.write(row, 7, 'ORDER')

  worksheet.repeat_rows(row)

  row += 1

  categories = [
    'Nuts • Nut Butters',
    'Seeds • Seed Butters',
    'Legumes',
    'Grains',
    'Dried Fruits',
    'Snacks • Trail Mix',
    'Sweeteners ',
    'Flavor Extracts',
    'Culinary Ingredients',
    'Mushrooms • Seaweeds',
    'Oils',
    'Vinegar • Miso • Tamari',
    'Culinary Herbs • Spices',
    'Spice Blends',
    'Botanical Herbs',
    'Teas • Tea Blends',
    'Nutrition Boosters',
    'Cacao • Cocoa',
    'Misc • Packaging']

  def write_row():

    worksheet.write(row, 0, code)
    worksheet.write(row, 1, certified_organic)
    worksheet.write(row, 2, name)

    worksheet.write(row, 3, origin)
    #if origin:
    #  worksheet.write_comment(row, 3, country_names[origin])

    worksheet.write(row, 5, price)

  name_size_re = re.compile('''
    ''' + not_numbers_or_letters + '''
    (''' + number + ''') # value
    \ *(?:th\ +)?
    ([a-z]*) # units
    \ *
    (''' + spaces_LETTERS + ''') # tail
    [^0-9a-z]*$
    ''', re.IGNORECASE|re.VERBOSE)
  description_head_re = re.compile('''
    Size:
    ''' + not_numbers_or_letters + '''
    (''' + spaces_LETTERS + ''') # head
    \ *
    (''' + number + ''') # value
    \ *(?:th\ +)?
    ([a-z]*) # units
    ''', re.IGNORECASE|re.VERBOSE)
  description_size_re = re.compile('''
    ''' + not_numbers_or_letters + '''
    (''' + spaces_LETTERS + ''') # head
    \ *
    (''' + number + ''') # value
    \ *(?:th\ +)?
    ([a-z]*) # units
    ''', re.IGNORECASE|re.VERBOSE)

  origins = set()
  for category in categories:

    worksheet.set_row(row, None, heading_category)
    worksheet.merge_range(row, 0, row, 7, category.upper())
    row += 1

    data[category].sort()
    for name, code, certified_organic, origin, description, price, skus, rules in data[category]:
      if rules:
        for size, code, price in rules:
          if not code:
            if size not in skus:
              continue

            code = skus[size]

          elif not size:
            for size in skus:
              if skus[size] == code:
                break

            else:
              continue

          worksheet.set_row(row, None, border)
          write_row()
          worksheet.write(row, 4, size)
          row += 1

        worksheet.set_row(row - 1, None, bottom)

        if origin:
          origins.update(origin.split('/'))

      elif skus:
        for size in skus:
          code = skus[size]

          worksheet.set_row(row, None, border)
          write_row()
          worksheet.write(row, 4, size)
          row += 1

        worksheet.set_row(row - 1, None, bottom)

        if origin:
          origins.update(origin.split('/'))

      elif 'specialorder' not in code.lower() and 'specialorder' not in description.replace(' ', '').lower():
        if code not in ['M41', 'M5200']:

          tail_match = name_size_re.search(name)
          if tail_match:
            parts = []

            start = tail_match.start()
            match = name_size_re.search(name, 0, start)

            while match and match.group(2):

              value, units, tail = match.groups('')
              parts.insert(0, fixup_groups(('', value, units, tail)))

              start = match.start()
              match = name_size_re.search(name, 0, start)

            value, units, tail = tail_match.groups('')
            size = fixup_groups(('', value, units, tail))

            if parts:
              if '.' not in value and '/' not in value and units not in ['ct', 'g', 'gal', 'gallon', 'kg', 'l', 'lb', 'lbs', 'ml', 'oz']:
                size += ' × ' + parts.pop(0)

              else:

                parts.append(size)
                size = parts.pop(0)

              if parts:
                size += ' (' + ' - '.join(parts) + ')'

              worksheet.write(row, 4, size)
              name = name[:start]

            elif units:

              worksheet.write(row, 4, size)
              name = name[:start]

          else:

            match = description_head_re.search(description)
            if match:
              parts = []

              head, value, units = match.groups('')
              match = description_size_re.match(description, match.end())

              while match and units:
                parts.append(fixup_groups((head, value, units, '')))

                head, value, units = match.groups('')
                match = description_size_re.match(description, match.end())

              size = fixup_groups((head, value, units, ''))
              if parts:
                if '.' not in value and '/' not in value and units not in ['ct', 'g', 'gal', 'gallon', 'kg', 'l', 'lb', 'lbs', 'ml', 'oz']:
                  size += ' × ' + parts.pop(0)

                else:

                  parts.append(size)
                  size = parts.pop(0)

                if parts:
                  size += ' (' + ' - '.join(parts) + ')'

                worksheet.write(row, 4, size)

              elif units:
                worksheet.write(row, 4, size)

        worksheet.set_row(row, None, bottom)
        write_row()
        row += 1

        if origin:
          origins.update(origin.split('/'))

  row += 1
  worksheet.merge_range(row, 0, row, 7, 'WHY ORGANIC MATTERS:', heading)
  row += 1

  worksheet.set_row(row, 42)
  worksheet.merge_range(row, 0, row, 7, 'For the health and wellbeing of people and the planet, our passion is to provide the highest quality and most nutritious foods, grown as locally, as fairly traded, and as sustainable as possible. We encourage bulk buying to minimize packaging and waste, keeping food simple and affordable.', text_wrap)
  row += 1

  row += 1
  worksheet.merge_range(row, 0, row, 7, 'ORDERING NOTES:', heading)
  row += 1

  worksheet.set_row(row, 162)
  worksheet.merge_range(row, 0, row, 7, None)
  worksheet.write_rich_string(row, 0, 'You can place your order ', bold, 'anytime', ' by phone ', bold, '250-505-2272', ' or email ', bold, 'info@omfoods.com', '. Out of stock items will not be placed on backorder.\nMinimum for ', bold, 'phone or email', ' orders is ', bold, '$150', ', ', bold, '$100', ' for ', bold, 'online', ' orders. Orders above $1,500 deduct 2%.\nOnce confirmed, you can pick up your order ', bold, 'Monday to Friday 10 AM to 5:30 PM', ' at our warehouse ', bold, '3505 Highway 6 • Nelson BC', '\nFor out of town orders, please enquire to receive a freight quote. ', bold, 'Only liquids in 16oz and 32oz containers will be shipped in glass.', '\nPayment is COD by cheque, e-transfer, or cash. ', bold, 'Returned cheques are subject to a $15 fee.\nPrices are subject to change without notice.', ' Pricing is mostly affected by currency exchange, transportation costs, and market fluctuations.\n', bold, 'RETURN POLICY:', ' We want you to be happy with our products! Organic foods such as grains, nuts, dried fruits, etc. are particularly susceptible to insect contamination. Please inspect products when they arrive and store them properly in cool, dark conditions. Organic Matters must be notified ', bold, 'within 7 days', ' if you wish to receive credit due to insect contamination or any other quality issues. Please return remaining products in their original packaging.', text_wrap)
  row += 1

  row += 1
  worksheet.merge_range(row, 0, row, 7, 'Thank you for eating the change you want to see!', thank_you)
  row += 1

  row += 1
  worksheet.merge_range(row, 0, row, 7, 'COUNTRY CODES:', heading)
  row += 1

  worksheet.set_row(row, 70)
  worksheet.merge_range(row, 0, row, 7, None)

  origins = list(origins)
  origins.sort()

  parts = []
  for origin in origins:
    parts.extend([bold, origin + ':', ' ' + country_names[origin] + '  '])

  parts.append(text_wrap)
  worksheet.write_rich_string(row, 0, *parts)

  row += 1

  worksheet.fit_to_pages(1, None)

  workbook.close()

  try:
    os.remove('CATALOG.xlsx')

  except OSError:
    pass

  os.symlink(filename + '.xlsx', 'CATALOG.xlsx')

  print 'Location: CATALOG.xlsx'
  print

else:

  print 'Content-Type: text/html'
  print
  print '<!DOCTYPE html>'
  print '<form enctype=multipart/form-data method=post>'
  print   '<input name=file type=file>'
  print   '<input type=submit>'
