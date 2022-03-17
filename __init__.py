"""
Description
-----------
A python interface to simulation and lab evaluation.

Author
------
Charles Sharman

Conventions
-----------
Store specs and plots unitless

License
-------
Distributed under the GNU LESSER GENERAL PUBLIC LICENSE Version 3.
View LICENSE for details.
"""

from __future__ import print_function

import os
import numpy as np
import matplotlib.pyplot as plt
plt.ion()
try:
    import pptx
except:
    print('Warning: Can\'t find pptx')
    pptx = None
import copy

config = {'project': '.', 'corners': '', 'typical_corner': '', 'current_corner': ''}

# Local Functions

def _results_dir():
    global config
    return os.path.join(os.path.abspath(config['project']), 'results')

def _wildcard_expand(name, possibles, must=True):
    """
    Returns a subset of possibles that satisfy name
        must=True forces name to be in possibles, even without a wildcard
    """
    lname = name.split('*')
    if len(lname) == 1:
        if (not must) or (must and lname[0] in possibles):
            return lname
        else:
            return []
    if len(lname) == 2:
        retvals = list(filter(lambda x: x.startswith(lname[0]), possibles))
        return list(filter(lambda x: x.endswith(lname[1]), retvals))
    else:
        print('Warning: Can only process one wildcard')
        return []

def _strip_file(name):
    """
    Rudimentary file stripper. Currently,
        * Removes comments
        * Skips blank space
    """
    fp = open(name, 'r')
    retval = []
    for line in fp:
        # Strip Comment
        cindex = line.find('#')
        if cindex > -1:
            line = line[:cindex]
        line = line.strip()
        if line:
            retval.append(line)
    fp.close()
    return retval

def _extract_units(s):
    """
    Extracts the units from a label.
    """
    try:
        units = s[s.find('(')+1]
    except:
        units = ' '
    return unit_adj(1, units)

def unit_adj(number, units):
    """
    Returns a unit-adjusted number
    """
    uindex = 'fpnum kMGT'.find(units[0])
    if uindex > -1:
        m = 10**(-3*(uindex-5))
    elif units == '%':
        m = 100
    else:
        m = 1
    return number * m

# Waveform Functions

def wave(xs, ys):
    """
    Returns a wave with individual xs and ys vectors
    """
    return np.transpose((xs, ys))

def value(w, x):
    """
    Returns the linearly interpolated y for a given x.  Assumes x is
    always increasing or decreasing.
    """
    if w[1, 0] > w[0, 0]:
        return np.interp(x, w[:,0], w[:,1])
    else:
        return np.interp(x, w[::-1][:,0], w[::-1][:,1])

def clip(w, x1, x2):
    """
    Clips a wave from x1 to x2.  Interpolates the end-points
    """
    y1 = value(w, x1)
    y2 = value(w, x2)
    keeps = np.nonzero((w[:,0] > x1) & (w[:,0] < x2))
    return wave((x1, w[:,0][keeps], x2), (y1, w[:,1][keeps], y2))

def crosses(w, y, edge=0):
    """
    Returns the x for a given y
      edge is 1 (positive), -1 (negative), 0 (either)
    """
    greater = (w[:,1] > y).astype(np.int)
    xs = greater[1:] - greater[:-1]
    if edge > 0.5:
        xs = xs > 0
    elif edge < -0.5:
        xs = xs < 0
    else:
        xs = np.abs(xs) > 0
    indices = np.nonzero(xs.astype(np.int))
    twave = np.transpose((w[:,1], w[:,0]))
    retval = []
    for index in indices[0]:
        retval.append(value(twave[index:index+2], y))
    return np.array(retval)

# I/O

def _set_corner(corner=''):
    global config
    if corner=='':
        corner = config['current_corner']
    return corner

def _check_dir(corner):
    write_dir = os.path.join(_results_dir(), corner)
    if not os.path.isdir(write_dir):
        os.makedirs(write_dir)
    return write_dir

def _read_specs(corner=''):
    corner = _set_corner(corner)
    retval = {}
    try:
        fp = open(os.path.join(_results_dir(), corner, 'specs'), 'r')
    except:
        fp = None
    if fp:
        for line in fp:
            key, value = line.strip().split()
            retval[key] = float(value)
        fp.close()
    return retval

def read_spec(name, corner=''):
    corner = _set_corner(corner)
    try:
        retval = _read_specs(corner)[name]
    except:
        print('Warning: Can\'t find spec', name)
        retval = None
    return retval

def write_spec(value, name, corner=''):
    corner = _set_corner(corner)
    specs = _read_specs(corner)
    specs[name] = value
    write_dir = _check_dir(corner)
    fp = open(os.path.join(write_dir, 'specs'), 'w')
    for spec in specs:
        fp.write(spec + ' ' + str(specs[spec]) + '\n')
    fp.close()

def read_wave(name, corner=''):
    corner = _set_corner(corner)
    xs = []
    ys = []
    full_name = os.path.join(_results_dir(), corner, name)
    try:
        fp = open(full_name, 'r')
    except:
        print('Warning: Can\'t find', full_name)
        fp = None
    if fp:
        for line in fp:
            x, y = line.strip().split()
            xs.append(float(x))
            ys.append(float(y))
        fp.close()
    return wave(xs, ys)

def write_wave(w, name, corner=''):
    corner = _set_corner(corner)
    write_dir = _check_dir(corner)
    fp = open(os.path.join(write_dir, name), 'w')
    for count in range(len(w[:,0])):
        fp.write(str(w[count,0]) + ' ' + str(w[count,1]) + '\n')
    fp.close()

# Evaluation

def script(name, corners=''):
    global config
    
    if corners=='':
        corners = config['corners']
    lines = _strip_file(os.path.join(os.path.abspath(config['project']), 'scripts'))
    names = _wildcard_expand(name, lines, must=False)
    for corner in corners.split():
        config['current_corner'] = corner
        for name in names:
            fullname = os.path.join(os.path.abspath(config['project']), name + '.py')
            print('Simulating %s at corner %s.' % (name, corner))
            #execfile(fullname)
            exec(open(fullname, 'r').read(), globals(), locals())

# Results

def plot(name, corners='', max_labels=6):
    """
    Plots a waveform in the plots file
        max_labels limits the number of legends to show

    TODO:
    * Add subpaths
    """
    if corners=='':
        corners = config['corners']
    fp = open(os.path.join(os.path.abspath(config['project']), 'plots'), 'r')
    found = False
    for line in fp:
        if line.startswith(name) and line[len(name)] in [',', ' ']:
            found = True
            break
    fp.close()
    if not found:
        print('Warning: Can\'t find plot', name)
        return
    lsplit = line.strip().split(',')
    try:
        names, title, xlabel, ylabel = lsplit[:4]
    except:
        print('Warning: Improper plot file syntax', line)
        return
    if len(lsplit) > 4:
        mplcmds = ','.join(lsplit[4:]).split(';')
    else:
        mplcmds = []
    plt.clf()
    plt.title(title)
    plt.xlabel(xlabel)
    xunits = _extract_units(xlabel)
    plt.ylabel(ylabel)
    yunits = _extract_units(ylabel)
    lnames = names.split()
    corners = corners.split()
    for corner in corners:
        cdir = os.path.join(_results_dir(), corner)
        wnames = []
        if os.path.isdir(cdir):
            possibles = os.listdir(cdir)
            for lname in lnames:
                wnames = wnames + _wildcard_expand(lname, possibles)
        for wname in wnames:
            w = read_wave(wname, corner)
            plt.plot(w[:,0]*xunits, w[:,1]*yunits, label=corner)
    for mplcmd in mplcmds:
        eval('plt.' + mplcmd)
    if len(corners) <= max_labels:
        plt.legend()
    plt.show()

def specs(ttype='mtm', corners=''):
    """
    Returns a csv spec table. ttype is either mtm (min/typ/max) or all.

    TODO:
    * Add subpaths
    """
    global config

    if corners=='':
        corners = config['corners']
    tables = {}
    corners = corners.split()
    for corner in corners:
        tables[corner] = _read_specs(corner)
    typical_corner = config['typical_corner']
    if typical_corner == '':
        print('Warning: Typical corner unspecified--choosing', corners[0])
        typical_corner = corners[0]
    stable = ''
    fp = open(os.path.join(os.path.abspath(config['project']), 'specs'), 'r')
    for line in fp:
        line = line.strip()
        if line:
            if line.startswith('*'): # A Title
                stable = stable + line + '\n'
            else:
                try:
                    name, title, ds_min, ds_typ, ds_max, units = line.split(',')
                except:
                    print('Warning: Can\'t parse', line)
                    name = None
                if name:
                    styp = tables[typical_corner].get(name, '')
                    smin = styp
                    smax = styp
                    line = ''
                    for corner in corners:
                        value = tables[corner].get(name, '')
                        if value != '':
                            if smin == '' or value < smin:
                                smin = value
                            if smax == '' or value > smax:
                                smax = value
                            if ttype == 'all':
                                line = line + '%.4g,' % unit_adj(value, units)
                        elif ttype == 'all':
                            line = line + ','
                    if smin != '':
                        smin = unit_adj(smin, units)
                    if smax != '':
                        smax = unit_adj(smax, units)
                    res = '-'
                    if ds_min and smin != '' and smin < float(ds_min):
                        res = 'F'
                    if ds_max and smax != '' and smax > float(ds_max):
                        res = 'F'
                    if ds_typ and (ds_typ[0] == '<'):
                        ds_typ = ds_typ[1:]
                        if smax != '' and smax > float(ds_typ):
                            res = 'F'
                    if ds_typ and (ds_typ[0] == '>'):
                        ds_typ = ds_typ[1:]
                        if smin != '' and smin < float(ds_typ):
                            res = 'F'
                    if ttype == 'mtm':
                        if styp != '':
                            styp = '%.4g' % unit_adj(styp, units)
                        if smin != '':
                            smin = '%.4g' % smin
                        if smax != '':
                            smax = '%.4g' % smax
                        line = '%s,%s,%s,' % (smin, styp, smax)
                    stable = stable + '%s,%s,%s,%s,%s%s,%s\n' % (title, ds_min, ds_typ, ds_max, line, units, res)
    return stable

# Printing

def print_plots():
    pass

def print_schematics():
    pass

def print_desrev():
    """
    Prints specs, schematics, and plots to a pptx
    template format:
        Slide 1: Title
        Slide 2: Bullets
        Slide 3: Spec
        Slide 4: Image
    """
    def _add_row(tbl):
        # Dangerous to use lower-level routines
        row = tbl._tbl.add_tr(table.rows[0].height)
        #index = len(tbl.rows)
        #cp_cell = table.cell(0, 0)
        #cp_cell = tbl._tbl.tc(0, 0)
        for col in range(9):
            cell = row.add_tc()
            #cell.get_or_add_tcPr(cp_cell.tcPr)
        """
        for col in range(9):
            cell = table.cell(col, index)
            #cell.fill = copy(cp_cell.fill)
            cell.margin_left = cp_cell.margin_left
            cell.margin_right = cp_cell.margin_right
            cell.margin_top = cp_cell.margin_top
            cell.margin_bottom = cp_cell.margin_bottom
            cell.vertical_anchor = cp_cell.vertical_anchor
        """
        
    #print_specs()
    #print_schematics()
    #print_plots()
    # Initialize
    #prs = pptx.Presentation()
    prs = pptx.Presentation('../template.pptx')
    write_dir = os.path.join(os.path.abspath(config['project']), 'reports')
    if not os.path.isdir(write_dir):
        os.makedirs(write_dir)

    # Specs
    slide = prs.slides[2] # Specs should be on slide 2
    table = slide.shapes[1].table # The table should be the second shape
    y = len(table.rows)
    rows_per_slide = 7 # Excludes title
    specs = print_specs()
    row_specs = specs.split('\n')
    num_slides = int(np.ceil(float(len(row_specs)) / rows_per_slide))
    #print('num_slides', num_slides)
    #row_title = 'Spec,Spec Min,Spec Typ,Spec Max,Res Min,Res Typ,Res Max,Units,Flag'
    for index_slide in range(num_slides):
        #slide = prs.slides.add_slide(prs.slide_layouts[5]) # title only
        if index_slide == num_slides-1: # last slide
            num_rows = len(row_specs) % rows_per_slide
        else:
            num_rows = rows_per_slide
        #shape = slide.shapes.add_table(rows=num_rows+1, cols=9)
        slide_rows = row_specs[rows_per_slide*index_slide:rows_per_slide*(index_slide+1)]
        #slide_rows.insert(0, row_title)
        for contents in slide_rows:
            _add_row(table)
            for x, content in enumerate(contents.split(',')):
                cell = table.cell(y, x)
                cell.text = content
            y = y + 1

    # Schematics

    # Plots
    prs.save(os.path.join(write_dir, 'desrev.pptx'))
