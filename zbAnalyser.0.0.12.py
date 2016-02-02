#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import os
from enum import Enum
import openpyxl
from openpyxl.worksheet import *
import copy


class Severity(Enum):
    """The values of column Severity"""
    Critical = 'Critical'
    Major = 'Major'
    Minor = 'Minor'
    Warning = 'Warning'
    Ok = 'Ok'

    def __str__(self):
        return self.value


class Check(Enum):
    """Header of checks"""
    Caption = 0
    Command = 1
    Output = 2
    Element = 3
    AlarmsReference = 4


class Alarm(Enum):
    """Header of alarms"""
    specificProblem = 0
    eventType = 1
    probableCause = 2
    perceivedSeverity = 3
    managedObject = 4
    RNCNodeType = 5


class ZbCheckRow():
    """ строка страницы файла Support Report """
    def __init__(self, checkname, order, severity=Severity.Ok, observation='No alarms', dateof='', nodename=''):
        super(ZbCheckRow, self).__init__()
        self.CheckName = checkname
        self.Severity = severity
        self.Observation = observation
        self.Order = order
        self.DateOf = dateof
        self.NodeName = nodename
        self.alarmsDetail = []
        self.alarmsCritical = 0
        self.alarmsMajor = 0
        self.alarmsMinor = 0
        self.alarmsWarning = 0
        self.alarmsTotal = 0
        self.alarmsCollision = 0

    def __str__(self):
        return '\t'.join([str(self.Order), self.CheckName, str(self.Severity), self.Observation])


def copy_rows(self, row_idx, cnt, above=False, copy_style=True, fill_formulae=True):
    """Inserts new (empty) rows into worksheet at specified row index.
    :param self: itself
    :param row_idx: Row index specifying where to insert new rows.
    :param cnt: Number of rows to insert.
    :param above: Set True to insert rows above specified row index.
    :param copy_style: Set True if new rows should copy style of immediately above row.
    :param fill_formulae: Set True if new rows should take on formula from immediately above row, filled with references new to rows.
    Usage:
    * insert_rows(2, 10, above=True, copy_style=False)
    """
    CELL_RE = re.compile("(?P<col>\$?[A-Z]+)(?P<row>\$?\d+)")
    row_idx = row_idx - 1 if above else row_idx

    def replace(m):
        row = m.group('row')
        prefix = "$" if row.find("$") != -1 else ""
        row = int(row.replace("$",""))
        row += cnt if row > row_idx else 0
        return m.group('col') + prefix + str(row)
    # First, we shift all cells down cnt rows...
    old_cells = set()
    old_fas = set()
    new_cells = dict()
    new_fas = dict()
    for c in self._cells.values():
        old_coor = c.coordinate
        # Shift all references to anything below row_idx
        if c.data_type == Cell.TYPE_FORMULA:
            c.value = CELL_RE.sub(
                replace,
                c.value
            )
            # Here, we need to properly update the formula references to reflect new row indices
            if old_coor in self.formula_attributes and 'ref' in self.formula_attributes[old_coor]:
                self.formula_attributes[old_coor]['ref'] = CELL_RE.sub(
                    replace,
                    self.formula_attributes[old_coor]['ref']
                )
        # Do the magic to set up our actual shift    
        if c.row > row_idx:
            old_coor = c.coordinate
            old_cells.add((c.row,c.col_idx))
            c.row += cnt
            new_cells[(c.row,c.col_idx)] = c
            if old_coor in self.formula_attributes:
                old_fas.add(old_coor)
                fa = self.formula_attributes[old_coor].copy()
                new_fas[c.coordinate] = fa
    for coor in old_cells:
        del self._cells[coor]
    self._cells.update(new_cells)
    for fa in old_fas:
        del self.formula_attributes[fa]
    self.formula_attributes.update(new_fas)
    # Next, we need to shift all the Row Dimensions below our new rows down by cnt...
    for row in range(len(self.row_dimensions)-1+cnt,row_idx+cnt,-1):
        new_rd = copy.copy(self.row_dimensions[row-cnt])
        new_rd.index = row
        self.row_dimensions[row] = new_rd
        del self.row_dimensions[row-cnt]
    # Now, create our new rows, with all the pretty cells
    row_idx += 1
    for row in range(row_idx,row_idx+cnt):
        # Create a Row Dimension for our new row
        new_rd = copy.copy(self.row_dimensions[row-1])
        new_rd.index = row
        self.row_dimensions[row] = new_rd
        for col in range(1,self.max_column):
            col = get_column_letter(col)
            cell = self.cell('%s%d'%(col,row))
            cell.value = self.cell('%s%d'%(col,row-1)).value
            source = self.cell('%s%d'%(col,row-1))
            if copy_style:
                cell.number_format = source.number_format
                cell.font        = source.font.copy()
                cell.alignment = source.alignment.copy()
                cell.border    = source.border.copy()
                cell.fill        = source.fill.copy()
            if fill_formulae and source.data_type == Cell.TYPE_FORMULA:
                s_coor = source.coordinate
                if s_coor in self.formula_attributes and 'ref' not in self.formula_attributes[s_coor]:
                    fa = self.formula_attributes[s_coor].copy()
                    self.formula_attributes[cell.coordinate] = fa
                # print("Copying formula from cell %s%d to %s%d"%(col,row-1,col,row))
                cell.value = re.sub(
                    "(\$?[A-Z]{1,3}\$?)%d"%(row - 1),
                    lambda m: m.group(1) + str(row),
                    source.value
                )     
                cell.data_type = Cell.TYPE_FORMULA
    # Check for Merged Cell Ranges that need to be expanded to contain new cells
    for cr_idx, cr in enumerate(self.merged_cell_ranges):
        self.merged_cell_ranges[cr_idx] = CELL_RE.sub(
            replace,
            cr
        )
Worksheet.copy_rows = copy_rows


class ZbAnalyser():
    """zbAnalyser! И этим всё сказано"""
    def __init__(self):
        super(ZbAnalyser, self).__init__()
        self.currentTemplate = '160123_t.xlsx'
        self.referenceError = 'Alarms_and_events.xlsx'
        self.dirs = { 'inputDir': './input', 'outputDir': './output', 'logDir': './log' }
        self.checks = (('Check active Alarms', 'alt', r'(?si)={10,}\nDate & Time \(Local\) +S +Specific Problem +MO ' +
                        '\(Cause/AdditionalInfo\)\n={10,}\n(.*?)\n?>>> Total: \d+ Alarms \(\d+ Critical, \d+ Major\)',
                        r'20\d{2}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} (\w) ((?:\w+ ?)+) +(.*)', 'Alarms_and_events.xlsx'),
                       ('Health check scheduler', 'get ManagedElement=1 healthCheckResult\|healthCheckSchedule',
                        r'(?si)={10,}\nMO +Attribute +Value\n={10,}\n(.*?)\n?={10,}\nTotal: \d+ Mos',
                        r'ManagedElement=\d+ +healthCheckSchedule t\[(\d+)\].*\n?(?: >>> Struct\[\d\] +has \d+.*)?\n?' +
                        '(?: >>> 1[.]time = \d{2}:\d{2})?\n?(?: >>> 2[.]weekday = \d+ \(\w+\))?', ''),
                       ('Check Event and System Logs', 'lgesmr 7d', r'(?s)={10,}\nTimestamp \(UTC\) +Type +Merged Log' +
                        ' Entry\n={10,}\n(.*)', r'(?i)[\d-]+ [\d:]+ +\w+ +(?:(?:(?:\w+=\w+),)+(\w+=\w+)|(?:Crash on (\d+), device=(\d+) [\w\d]+)) +(.+)', ''))
        self.output = []
        self.wb = None
        self.log = None
        self.alarms = None
        self.alarmsReferenceName = ''
        self.logdate = None
        if not os.path.exists(self.dirs['inputDir']):
            os.mkdir(self.dirs['inputDir'])
        if not os.path.exists(self.dirs['outputDir']):
            os.mkdir(self.dirs['outputDir'])
        if not os.path.exists(self.dirs['logDir']):
            os.mkdir(self.dirs['logDir'])

    def init_alarms(self):
        wb = openpyxl.load_workbook(filename=self.referenceError)
        ws = wb['Alarms']
        self.alarms = []
        headerfounded = False
        for row in ws.iter_rows():
            if headerfounded:
                b = []
                for cell in row:
                    if cell.column in ('A', 'B', 'C', 'D', 'E', 'F'):
                        b.append(cell.value.strip(' '))
                if b[0] is not None:
                    self.alarms.append(tuple(b))
            if row[0].value and row[0].value.lower().strip(' ') == 'specificproblem':
                headerfounded = True
        self.alarms = tuple(self.alarms)
        return self.alarms

    def parseLog(self, nodename):
        if self.log is None:
            print('No log!')
            return
        logdatere = re.search(r'Logging to file [/\w\d]+/(\d{4}-\d{2}-\d{2})', self.log)
        if logdatere:
            self.logdate = logdatere.group(1)
        for num, check in enumerate(self.checks):
            nextStr = ZbCheckRow(checkname=check[Check.Caption.value], order=num, nodename=nodename)
            commandRegExp = r'(?m)^(?:[\w\d.]+)> %s\n((?:.*\n?(?!(?:^[\w\d.]+)>))*)'
            if (check[Check.AlarmsReference.value] != '' and
                os.path.exists(check[Check.AlarmsReference.value]) and
               self.alarmsReferenceName != check[Check.AlarmsReference.value]):
                self.alarmsReferenceName = check[Check.AlarmsReference.value]
                self.init_alarms()
            outputRE = re.search(commandRegExp % check[Check.Command.value], self.log)
            if outputRE is None:
                print('outputRE is fail!')
                continue
            output = outputRE.group(1)
            if output is None:
                print('Command RegExp fail')
                continue
            commandDateRE = re.search(r'(\d{6})-\d{2}:\d{2}:\d{2}', output)
            if commandDateRE:
                nextStr.DateOf = commandDateRE.group(1)
            outputLinesRE = re.search(check[Check.Output.value], output)
            if outputLinesRE is None:
                print('outputLinesRE is fail!')
                continue
            outputLines = outputLinesRE.group(1)
            elementRE = re.compile(check[Check.Element.value])
            if check[Check.Command.value] in [self.checks[0][Check.Command.value]]:
                if elementRE.search(outputLines):
                    for element in elementRE.findall(outputLines):
                        if self.alarmsReferenceName == check[Check.AlarmsReference.value] and self.alarms is not None:
                            for alarm in self.alarms:
                                if element[1].lower().strip(' ') == alarm[Alarm.specificProblem.value].lower():
                                    if element[0] == 'c' and alarm[Alarm.perceivedSeverity.value].lower() == 'critical':
                                        nextStr.alarmsCritical += 1
                                        nextStr.alarmsDetail.append(element[1])
                                    elif element[0] == 'M' and alarm[Alarm.perceivedSeverity.value].lower() == 'major':
                                        nextStr.alarmsMajor += 1
                                        nextStr.alarmsDetail.append(element[1])
                                    elif element[0] == 'm' and alarm[Alarm.perceivedSeverity.value].lower() == 'minor':
                                        nextStr.alarmsMinor += 1
                                    elif element[0] == 'w' and alarm[Alarm.perceivedSeverity.value].lower() == 'warning':
                                        nextStr.alarmsWarning += 1
                                    else:
                                        nextStr.alarmsCollision += 1
                                        print('Unknown perceivedSeverity!')
                                    print(element)
                                    break
                nextStr.alarmsTotal = nextStr.alarmsCritical + nextStr.alarmsMajor + nextStr.alarmsMinor + \
                                      nextStr.alarmsWarning + nextStr.alarmsCollision
                if nextStr.alarmsTotal > 0:
                    nextStr.Observation = 'Total %d alarms:' % nextStr.alarmsTotal
                    if nextStr.alarmsCritical > 0:
                        nextStr.Observation += ' %d critical' % nextStr.alarmsCritical
                    if nextStr.alarmsMajor > 0:
                        nextStr.Observation += (',' if nextStr.alarmsCritical > 0 else '') + \
                                               ' %d major' % nextStr.alarmsMajor
                    if nextStr.alarmsMinor > 0:
                        nextStr.Observation += (',' if nextStr.alarmsCritical > 0 or nextStr.alarmsMajor > 0 else '') + \
                                               ' %d minor' % nextStr.alarmsMinor
                    if nextStr.alarmsWarning > 0:
                        nextStr.Observation += (',' if nextStr.alarmsCritical > 0 or nextStr.alarmsMajor > 0 or nextStr.alarmsMinor > 0 else '') + \
                                               ' %d warning' % nextStr.alarmsWarning
                    if nextStr.alarmsCollision > 0:
                        nextStr.Observation += (',' if nextStr.alarmsCritical > 0 or nextStr.alarmsMajor > 0 or nextStr.alarmsMinor > 0 or nextStr.alarmsWarning > 0 else '') + \
                                               ' %d collision' % nextStr.alarmsCollision
                    if nextStr.alarmsCritical > 0:
                        nextStr.Severity = Severity.Critical
                    elif nextStr.alarmsMajor > 0:
                        nextStr.Severity = Severity.Major
                    elif nextStr.alarmsMinor > 0:
                        nextStr.Severity = Severity.Minor
                    elif nextStr.alarmsWarning > 0:
                        nextStr.Severity = Severity.Warning
            if check[Check.Command.value] == self.checks[1][Check.Command.value]:
                element = elementRE.search(outputLines)
                if element is None or element.groups()[0] == '0':
                    nextStr.Severity = Severity.Warning
                    nextStr.Observation = 'Health Check Schedule is NOK'
                else:
                    nextStr.Observation = 'Health Check Schedule is OK'
            if check[Check.Command.value] == self.checks[2][Check.Command.value]:
                if elementRE.search(outputLines):
                    sum = 0
                    MOs = set()
                    nextStr.Observation = ''
                    for element in elementRE.findall(outputLines):
                        if element[3] is not None and element[3].lower().find('ranap_cninitiatedresetresource') >= 0:
                            nextStr.Severity = Severity.Critical
                            MOs.add(element[0])
                            sum += 1
                    if sum != 0:
                        if nextStr.Observation != '':
                            nextStr.Observation += '\n'
                        nextStr.Observation += 'Ranap_CNInitiatedResetResource %s sum: %d' % (str(MOs), sum)
                    sum = 0
                    MOs = set()
                    for element in elementRE.findall(outputLines):
                        if element[3] is not None and element[3].lower().find('ipethpacketdatarouter_cnnotrespondingtogtpecho') >= 0:
                            if nextStr.Severity != Severity.Critical:
                                nextStr.Severity = Severity.Major
                            MOs.add(element[0])
                            sum += 1
                    if sum != 0:
                        if nextStr.Observation != '':
                            nextStr.Observation += '\n'
                        nextStr.Observation += 'IpEthPacketDataRouter_CnNotRespondingToGTPEcho %s sum: %d' % (str(MOs), sum)
                    sum = 0
                    MOs = set()
                    prevdevice = ''
                    for element in elementRE.findall(outputLines):
                        if element[1] is None or element[1] == '':
                            continue
                        if prevdevice != '' and prevdevice != element[2]:
                            nextStr.Severity = Severity.Critical
                        elif int(element[1]) > 1 and nextStr.Severity != Severity.Critical:
                            nextStr.Severity = Severity.Major
                        elif int(element[1]) == 1 and nextStr.Severity != Severity.Critical and nextStr.Severity != Severity.Major:
                            nextStr.Severity = Severity.Minor
                        MOs.add(', '.join(['Crash on %s' % element[1], 'device=%s' % element[2]]))
                        sum += 1
                    if sum != 0:
                        if nextStr.Observation != '':
                            nextStr.Observation += '\n'
                        nextStr.Observation += '%s sum: %d' % (str(MOs), sum)
                    if nextStr.Observation == '':
                        nextStr.Observation = 'No alarms'
            self.output.append(nextStr)

    def uparse(self, order):
        if self.inStr is None: return None
        critical = 0
        major = 0
        minor = 0
        warning = 0
        for row in self.alarms:
            if self.inStr.lower().find(str(row[0]).lower()) != -1:
                if row[3].lower() == 'critical':
                    critical += 1
                elif row[3].lower() == 'major':
                    major += 1
                elif row[3].lower() == 'minor':
                    minor += 1
                else:
                    warning += 1
        Observation = '%i Alarms:' % (critical+major+minor+warning)
        Observation += ' %i Critical' % critical if critical > 0 else ''
        Observation += ' %i Major' % major if major > 0 else ''
        Observation += ' %i Minor' % minor if minor > 0 else ''
        Observation += ' %i Warning' % warning if warning > 0 else ''
        if critical > 0:
            S = Severity.Critical
        elif major > 0:
            S = Severity.Major
        elif minor > 0:
            S = Severity.Minor
        elif warning > 0:
            S = Severity.Warning
        else:
            S = Severity.Ok
            Observation = 'Log is clear'
        return ZbCheckRow(self.checkNames[self.command], S, Observation, order, None)

    def savexls(self,filename):
        count = 1
        output = os.path.join(self.dirs['outputDir'], filename+'.xlsx')
        while os.path.exists(output):
            output = os.path.join(self.dirs['outputDir'], filename+str(count)+'.xlsx')
            count += 1
        self.wb.save(output)
        return output

    def writexls(self,filename):
        fs_init_row = 6
        self.wb = openpyxl.load_workbook(filename = self.currentTemplate)
        fs = self.wb['Front Sheet']
        for cell in fs._cells.values():
            if cell.comment == 'AutoCopy':
                fs_init_row = cell.row
                cell.comment = None
                break
        file_number = len(os.listdir(self.dirs['inputDir']))
        if file_number > 2:
            fs.copy_rows(fs_init_row, file_number-2, above=False, copy_style=True, fill_formulae=True)
        else:
            print('Is need more than 2 log files!')
            return
        tmpl = self.wb['Controller log template']
        for num, inFile in enumerate(os.listdir(self.dirs['inputDir'])):
            ws = copy.copy(tmpl)
            ws.title = inFile
            self.wb._add_sheet(ws)
        output = self.savexls(filename)
        self.wb = openpyxl.load_workbook(filename=output)
        try:
            fs = self.wb['Front Sheet']
            tmpl = self.wb['Controller log template']
            for num, inFile in enumerate(os.listdir(self.dirs['inputDir'])):
                print(inFile)
                ws = self.wb[inFile]
                with open(os.path.join(self.dirs['inputDir'], inFile), 'r') as f:
                    self.log = f.read()
                self.output = []
                self.parseLog(inFile)
                for cell in ws.rows[0]:
                    cell.value = cell.value.replace('v<#LogDate#>', self.logdate) if cell.value else None
                for row in self.output:
                    cur_row = int(row.Order)+5
                    ws.copy_rows(cur_row, 1, above=False, copy_style=True, fill_formulae=True)
                    for cell in ws.rows[cur_row-1]:
                        cell.value = cell.value.replace('v<#CheckName#>', row.CheckName) if cell.value else None
                        cell.value = cell.value.replace('v<#Severity#>', str(row.Severity)) if cell.value else None
                        cell.value = cell.value.replace('v<#Observation#>', row.Observation) if cell.value else None
                        cell.value = cell.value.replace('v<#DateOf#>', row.DateOf) if cell.value else None
                for cell in ws.rows[cur_row]:
                    cell.value = ''
                cur_row = num+fs_init_row
                for cell in fs.rows[cur_row-1]:
                    cell.value = cell.value.replace('v<#FileName#>', inFile) if cell.value else None
                    cell.value = cell.value.replace('v<#MaxRow#>', str(ws.max_row)) if cell.value else None
                    formula = re.compile(r'f<#(.*)#>')
                    if formula.search(str(cell.value)):
                        cell.value = formula.sub(r'\1', cell.value).replace(';',',')
                        cell.data_type = Cell.TYPE_FORMULA
            if tmpl: self.wb.remove_sheet(tmpl)
        # except Exception, e:
            # raise e
        # else:
            # pass
        finally:
            self.wb.save(output)


def main():
    zloyB = ZbAnalyser()
    # zloyB.init_alarms()
    # for row in zloyB.alarms:
        # print(row)
    zloyB.writexls('Preemptive_Support_Report_WRAN_Vimpelcom_NorthWest_')

    # BSC => GRAN
    # RNC => WRAN
    # zloyB.test('test')

if __name__ == '__main__':
    main()
