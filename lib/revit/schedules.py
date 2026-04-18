# -*- coding: utf-8 -*-
import os
import re
from Autodesk.Revit.DB import ViewSchedule, SectionType, FilteredElementCollector


def sanitize_filename(name, max_length=80):
    name = unicode(name) if not isinstance(name, unicode) else name
    clean = re.sub(r'[\\/*?:"<>|]', u'_', name)
    return clean[:max_length].strip() or u'planilla'


def get_non_template_schedules(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    return [s for s in collector if not s.IsTemplate]


def get_schedule_data(schedule):
    data = []
    body = schedule.GetTableData().GetSectionData(SectionType.Body)
    for r in range(body.NumberOfRows):
        row = []
        for c in range(body.NumberOfColumns):
            value = schedule.GetCellText(SectionType.Body, r, c) or u''
            if not isinstance(value, unicode):
                value = unicode(value)
            row.append(value)
        if any(cell.strip() for cell in row):
            data.append(row)
    return data


def _escape_csv_cell(value):
    value = value.replace(u'"', u'""')
    if u';' in value or u'\n' in value or u'\r' in value or u'"' in value:
        return u'"{}"'.format(value)
    return value


def export_schedule_to_csv(schedule, folder):
    if not os.path.exists(folder):
        os.makedirs(folder)
    filename = sanitize_filename(schedule.Name) + u'.csv'
    filepath = os.path.join(folder, filename)
    rows = get_schedule_data(schedule)
    with open(filepath, 'wb') as f:
        for row in rows:
            safe_row = [_escape_csv_cell(cell) for cell in row]
            line = u';'.join(safe_row) + u'\r\n'
            f.write(line.encode('utf-8'))
    return filepath
