
"""
Parse all the Gaia-ESO Survey run reports and calculate integrated time for 
the Milky Way, Cluster, and Calibrations.
"""

import time
from glob import glob

import numpy as np

import xlrd



reports = glob("reports/*.xl*") # (.xls, .xlt, .xlsx)


def extract_observations(filename):

    book = xlrd.open_workbook(filename)

    sheet_names = book.sheet_names()
    sheet_indices = range(len(sheet_names))

    rows = []
    for sheet_index, sheet_name in zip(sheet_indices, sheet_names):

        """
        if len(sheet_indices) > 1 and sheet_name.startswith("Sheet"):
            print("Skipping sheet name '{0}' in {1} (index {2})".format(
                sheet_name, filename, sheet_index))
            continue
        """

        sheet = book.sheets()[sheet_index]
        n_rows = sheet.nrows
        if n_rows == 0: continue

        print("Doing sheet {0} {1} on {2}".format(sheet_index, sheet_name,
            filename))

        # Find the row where the first column starts with 'OBID'
        # This should be row number 14, but let's be sure.

        skip_rows = 0
        for i in range(n_rows):

            values = [unicode(_.value).strip().lower() for _ in sheet.row(i)]
            if "obid" in values:
                skip_rows = i + 1
                break

        else:
            #if filename == "reports/SPS_ObsRunReport_193.B-0936(C).xls":
            #    raise a
            print("Warning: could not find header row on sheet {}".format(sheet_name))
            
        indices = [0, 1, 2, 3, 4, 4]
        col_names = ("obid", "ob name", "ob start time", "ob end time", "exptime")
        header = [str(_.value).strip().lower() for _ in sheet.row(skip_rows - 1)]
        for j, name in enumerate(col_names):
            try:
                indices[j] = header.index(name)
            except:
                print("Couldn't find {0} in sheet {1} of filename {2}".format(
                    name, sheet_name, filename))
                continue

        for i in range(skip_rows, n_rows):

            if len(indices) > len(sheet.row(i)):
                #print("Skipping row: {}".format(sheet.row(i)))
                continue            

            OB_id = unicode(sheet.row(i)[indices[0]].value).strip()
            OB_name = unicode(sheet.row(i)[indices[1]].value).strip()
            
            if sheet.row(i)[indices[2]].ctype == 3 \
            and not sheet.row(i)[indices[2]].value > 40000:
                foo = xlrd.xldate_as_tuple(sheet.row(i)[indices[2]].value, 0)
                ut_start = "{0:02d}:{1:02d}".format(foo[-3], foo[-2])

            elif sheet.row(i)[indices[2]].ctype == 3 \
            and sheet.row(i)[indices[2]].value > 40000:
                foo = xlrd.xldate_as_tuple(sheet.row(i)[indices[2]].value, 0)
                ut_start = "{0}-{1:02d}-{2:02d}T{3:02d}:{4:02d}".format(*foo)
            else:
                ut_start = unicode(sheet.row(i)[indices[2]].value).strip()

            if sheet.row(i)[indices[3]].ctype == 3 \
            and not sheet.row(i)[indices[3]].value > 40000:
                foo = xlrd.xldate_as_tuple(sheet.row(i)[indices[3]].value, 0)
                ut_end = "{0:02d}:{1:02d}".format(foo[-3], foo[-2])
            
            elif sheet.row(i)[indices[3]].ctype == 3 \
            and sheet.row(i)[indices[3]].value > 40000:
                foo = xlrd.xldate_as_tuple(sheet.row(i)[indices[3]].value, 0)
                ut_end = "{0}-{1:02d}-{2:02d}T{3:02d}:{4:02d}".format(*foo)
            else:
                ut_end = unicode(sheet.row(i)[indices[3]].value).strip()
            

            exptime = unicode(sheet.row(i)[indices[4]].value).strip()
            rows.append({
                "report": filename,
                "OBID": OB_id,
                "OB Name": OB_name,
                "OB Start time": ut_start,
                "OB End time": ut_end,
                "EXPTIME": exptime
            })

    return rows




rows = []
for filename in reports:
    print(filename)
    rows.extend(extract_observations(filename))

bad_rows = [row for row in rows if not row["OBID"].startswith("200")]
good_rows = [row for row in rows if row["OBID"].startswith("200")]


# Fix the time stamps on good rows
for row in good_rows:
    if "/2012" in row["OB Start time"]:
        date, stime = row["OB Start time"].split(" ")
        month, day, year = date.split("/")
        day, month, year = date.split("/")
        row["OB Start time"] = "{0}-{1:02d}-{2:02d}T{3}".format(year,
            int(month), int(day), stime)

    if row["OB Start time"].count(" ") > 0 and len(row["OB Start time"]) < 10:
        row["OB Start time"] = row["OB Start time"].replace(" ", ":")[:5]

    #if len(row["OB Start time"]) == 8 and row["OB Start time"].count(":") == 2:
    #    row["OB Start time"] = row["OB Start time"][:-3]

    row["OB Start time"] = row["OB Start time"][:19]


    row['OB Start time'] = row["OB Start time"].replace("-4T00:", "-04T00:")
    

    if row["OB Start time"].count(":") == 1:
        row["OB Start time"] += ":00"


    if "/2012" in row["OB End time"]:
        date, etime = row["OB End time"].split(" ")
        month, day, year = date.split("/")
        day, month, year = date.split("/")
        row["OB End time"] = "{0}-{1:02d}-{2:02d}T{3}".format(year, int(month), int(day), etime)


    row['OB End time'] = row["OB End time"].replace("-4T00:", "-04T00:")

    if row["OB End time"].count(" ") > 0 and len(row["OB End time"]) < 10:
        row["OB End time"] = row["OB End time"].replace(" ", ":")[:5]


    #if len(row["OB End time"]) == 8 and row["OB End time"].count(":") == 2:
    #    row["OB End time"] = row["OB End time"][:-3]

    if row["OB End time"].count(":") == 1:
        row["OB End time"] += ":00"

    row["OB End time"] = row["OB End time"][:19]

    # Clean up the timestamps that have exposure times as end times.
    if len(row["OB End time"]) == 3:


        stime = row["OB Start time"].split("T")[1]
        hr, minutes, secs = stime.split(":")

        end_secs = int(secs) + int(float(row["OB End time"]))
        if end_secs > 59:
            minutes = int(minutes) + 1
            if minutes > 59:
                raise FUCKOFF
        else:
            minutes = int(minutes)

        row["OB End time"] = row["OB Start time"].split("T")[0] + "T{0}:{1:02d}:{2:02d}".format(
            hr, minutes, end_secs)


# Clean up times without dates.
for row in good_rows:
    slen = len(row["OB Start time"])
    elen = len(row["OB End time"])

    if slen == 19 and elen == 19:
        continue

    if (slen == 8 and elen == 19) or (elen == 8 and slen == 19):
        raise FUCKOFF

    if slen == 0 or elen == 0:
        continue

    assert slen == elen
    assert slen == 8

    shr, smin, ssec = row["OB Start time"].split(":")

    row["OB Start time"] = "2000-01-01T{0}:{1}:{2}".format(shr, smin, ssec)

    ehr, emin, esec = row["OB End time"].split(":")

    shr = int(shr)
    ehr = int(ehr)

    if ehr < shr: # the OB passed midnight
        row["OB End time"] = "2000-01-02T{0:02d}:{1}:{2}".format(ehr, emin, esec)
    else:
        row["OB End time"] = "2000-01-01T{0:02d}:{1}:{2}".format(ehr, emin, esec)


# Calculate exposure times for all objects, based on the time stamps.
for row in good_rows:
    slen = len(row["OB Start time"])
    elen = len(row["OB End time"])

    if slen == 19 and elen == 19:
        format = '%Y-%m-%dT%H:%M:%S'
        start = time.strptime(row["OB Start time"], format)
        end = time.strptime(row["OB End time"], format)

        row["EXPTIME"] = time.mktime(end) - time.mktime(start)
        if row["EXPTIME"] < 0:
            print("Warning: Found an OB with a negative exposure time")
            print(row["OB Start time"], row["OB End time"])
            row["EXPTIME"] = np.abs(row["EXPTIME"])
            row["EXPTIME"] = np.nan
    else:
        row["EXPTIME"] = np.nan

    row["ASSIGNED"] = None

# Allocate OBs to either Clusters, Calibrations, or MW.
unassigned = []
for row in good_rows:
    name = row["OB Name"].upper()

    if name.upper().startswith(("GES_MW_", "GES_CRT_")) or "MW_" in name \
    or "COROT" in name.upper() or "GES_MW" in name or "BULGE" in name:
        assigned = "MW"

    elif name.upper().startswith("GES_CL") or "GES_CL" in name or "GES_CI" in name:
        assigned = "CL"

    elif name.startswith("GES_CAL"):
        assigned = "CALIBRATIONS"

    else:
        assigned = "UNKNOWN"
        print(name)
        unassigned.append(name)

    row["ASSIGNED"] = assigned

# Calculate total integrated time for each.

MW = [row["EXPTIME"] for row in good_rows if row["ASSIGNED"] == "MW"]
CL = [row["EXPTIME"] for row in good_rows if row["ASSIGNED"] == "CL"]
CALIBRATIONS = [row["EXPTIME"] for row in good_rows if row["ASSIGNED"] == "CALIBRATIONS"]
UNKNOWN = [row["EXPTIME"] for row in good_rows if row["ASSIGNED"] == "UNKNOWN"]

MW, CL, CALIBRATIONS, UNKNOWN = map(np.array, (MW, CL, CALIBRATIONS, UNKNOWN))


total_MW = np.nansum(MW)
total_CL = np.nansum(CL)
total_CALIBRATIONS = np.nansum(CALIBRATIONS)
total_UNKNOWN = np.nansum(UNKNOWN)

total = total_MW + total_CL + total_CALIBRATIONS + total_UNKNOWN

print(total_MW, total_CL, total_CALIBRATIONS, total_UNKNOWN)

print(total_MW/total, total_CL/total, total_CALIBRATIONS/total, total_UNKNOWN/total)
# Fraction.


