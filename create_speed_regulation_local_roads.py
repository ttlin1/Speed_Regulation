import arcpy
import os
import sys
from xlrd import open_workbook

sys.dont_write_bytecode = True
sys.path.append(r"C:\TOM_Scripts\MassDOT")
arcpy.env.overwriteOutput = True


def convert_excel_to_text(excel_file, geodatabase):
    """Convert the Excel spreadsheet to a geodatabase table"""
    work_book = open_workbook(excel_file, "rb")
    sheet = work_book.sheet_by_index(0)
    field_type_dict = {"REGULATION": "LONG",
                       "TOWN": "TEXT",
                       "ROUTENUM": "TEXT",
                       "DIRECTION": "TEXT",
                       "FROM_LOC": "TEXT",
                       "FROM_DISTANCE_FEET": "LONG",
                       "END_LOC": "TEXT",
                       "END_DISTANCE_FEET": "LONG",
                       "DISTANCE": "DOUBLE",
                       "SPEED": "SHORT",
                       "AMENDMENT": "TEXT",
                       "TIME_PERIOD": "TEXT",
                       "LINK": "TEXT"}

    arcpy.CreateTable_management(geodatabase, "new_table")
    new_table = os.path.join(geodatabase, "new_table")

    data_list = []
    for row_index in range(sheet.nrows):
        current_row = []
        for col_index in range(sheet.ncols):
            current_row.append(sheet.cell(row_index, col_index).value)
        if current_row[0] == "REGULATION":
            legend = current_row
        else:
            data_list.append(current_row)

    for l in legend:
        arcpy.AddField_management(new_table, l, field_type_dict[l])

    cursor = arcpy.da.InsertCursor(new_table, legend)
    for d in data_list:
        cursor.insertRow(d)
    del cursor
    return


def create_speed_regulation(excel_file_path, base_geodatabase, output_geodatabase, roads_fc, jurisdiction):
    """Automate the linear referencing of the speed regulation data"""
    municipal_boundary = os.path.join(base_geodatabase, "Municipal_Boundary")
    state_boundary = os.path.join(base_geodatabase, "Neighboring_States")

    # import excel table
    new_table = os.path.join(output_geodatabase, "new_table")
    if arcpy.Exists(new_table):
        arcpy.Delete_management(new_table)
    convert_excel_to_text(excel_file_path, output_geodatabase)

    town_route_num_dict = {}
    town_street_name_dict = {}

    # return feature layer for location/route
    def determine_loc_lyr(loc, twn, gdb):
        loc = str(loc)
        if loc.find(" TOWN LINE") > 0:
            town_name = loc.replace(" TOWN LINE", "")
            loc_lyr = arcpy.MakeFeatureLayer_management(municipal_boundary, town_name, "\"TOWN\" = '" + town_name + "'")
        elif loc.find(" STATE LINE") > 0:
            state_name = loc.replace(" STATE LINE", "")
            loc_lyr = arcpy.MakeFeatureLayer_management(state_boundary, state_name, "\"STATE_NAME\" = '" + state_name + "'")
        else:
            if loc in town_route_num_dict[twn]:
                temp = arcpy.MakeFeatureLayer_management("tempTownRoad", "temp", "(\"RT_NUMBER\" = '" + loc +
                                                         "' OR \"ALTRTNUM1\" = '" + loc + "' OR \"ALTRTNUM2\" = '" +
                                                         loc + "' OR \"ALTRTNUM3\" = '" + loc +
                                                         "' OR \"ALTRTNUM4\" = '" + loc + "')")
                loc_lyr = arcpy.CopyFeatures_management(temp, os.path.join(gdb, "Route" + loc.replace(" ", "_")))
            elif loc in town_street_name_dict[twn]:
                temp = arcpy.MakeFeatureLayer_management("tempTownRoad", "temp", "\"STREET_NAM\" = '" + loc + "'")
                loc_lyr = arcpy.CopyFeatures_management(temp, os.path.join(gdb, loc.replace(" ", "_").replace("-","_").replace("'", "")))
        return loc_lyr

    # calculate from mileage and to mileage for same direction of road
    arcpy.AddField_management(new_table, "START_MP", "DOUBLE")
    arcpy.AddField_management(new_table, "END_MP", "DOUBLE")
    arcpy.AddField_management(new_table, "Last_Segment", "TEXT")

    # temp field calculation
    arcpy.AddField_management(new_table, "JURISDICTION", "SHORT")
    arcpy.AddField_management(new_table, "NOTES", "TEXT")

    # jurisdiction
    arcpy.CalculateField_management(new_table, "JURISDICTION", jurisdiction, "PYTHON_9.3")

    current = ("", "", "", "", 0)
    mileage_total = 0
    rows = arcpy.UpdateCursor(new_table)
    last_segment_in_route_dict = {}
    for row in rows:
        key = (row.REGULATION, row.TOWN, row.ROUTENUM, row.DIRECTION, row.FROM_LOC,
               row.FROM_DISTANCE_FEET)
        if key != current:
            row.START_MP = round(row.FROM_DISTANCE_FEET / 5280.0, 2)
            row.END_MP = round(row.DISTANCE + row.FROM_DISTANCE_FEET / 5280.0, 2)
            current = key
            mileage_total = row.DISTANCE
        else:
            row.START_MP = round(mileage_total, 2)
            row.END_MP = round(mileage_total + row.DISTANCE, 2)
            mileage_total += row.DISTANCE
        last_segment_in_route_dict[(row.REGULATION, row.TOWN, row.ROUTENUM, row.DIRECTION, row.FROM_LOC, row.END_LOC)] = row.OBJECTID
        rows.updateRow(row)

    del rows

    rows = arcpy.UpdateCursor(new_table)
    for row in rows:
        if row.OBJECTID == last_segment_in_route_dict[(row.REGULATION, row.TOWN, row.ROUTENUM, row.DIRECTION,
                                                       row.FROM_LOC, row.END_LOC)]:
            row.Last_Segment = "Y"
            row.END_MP = round(row.END_MP - row.END_DISTANCE_FEET / 5280.0, 2)
        rows.updateRow(row)

    del rows

    # create dictionary of route ID to route data
    route_dict = {}
    route_list = []
    route_id_to_distance = {}
    for row in arcpy.SearchCursor(new_table):
        if (row.ROUTENUM, row.DIRECTION) not in route_dict:
            route_dict[(row.ROUTENUM, row.DIRECTION)] = [row.REGULATION, row.TOWN, row.ROUTENUM, row.DIRECTION,
                                                         row.FROM_LOC, row.FROM_DISTANCE_FEET, row.END_LOC,
                                                         row.END_DISTANCE_FEET]
        if (row.ROUTENUM, row.DIRECTION) not in route_list:
            route_list.append((row.ROUTENUM, row.DIRECTION))
        if (row.ROUTENUM, row.DIRECTION) not in route_id_to_distance:
            route_id_to_distance[(row.ROUTENUM, row.DIRECTION)] = row.DISTANCE
        else:
            route_id_to_distance[(row.ROUTENUM, row.DIRECTION)] += row.DISTANCE

    # calculate origin and destination points of speed regulation
    speed_fc_list = []
    for (route, direction) in route_list:
        print route
        regulation = route_dict[(route, direction)][0]
        town = route_dict[(route, direction)][1]
        route_number = str(route_dict[(route, direction)][2])
        from_location = str(route_dict[(route, direction)][4])
        from_distance = route_dict[(route, direction)][5] / 5280.0
        end_location = str(route_dict[(route, direction)][6])
        end_distance = route_dict[(route, direction)][7] / 5280.0
        temp_route = os.path.join(output_geodatabase, "tempRoute")
        origin_point = os.path.join(output_geodatabase, "origin")
        end_point = os.path.join(output_geodatabase, "end")
        calibrate_points = os.path.join(output_geodatabase, "calibrate_points")
        temp_routes = os.path.join(output_geodatabase, "temp_routes")
        calibrated_route = os.path.join(output_geodatabase, "calibrateRoute")
        speed_route = os.path.join(output_geodatabase, "speed" + str(route).replace(" ", "").replace("&", "").replace("/", "").replace("-", "").replace("'", "").replace(".", "") + direction)
        sort_fc = os.path.join(output_geodatabase, "sorted")

        # make lists of route numbers and street names for town
        if town not in town_route_num_dict:
            arcpy.MakeFeatureLayer_management(roads_fc, "tempTownRoad", "\"MGIS_TOWN\" = '"
                                              + town + "'")
            town_route_num_dict[town] = []
            town_street_name_dict[town] = []
            for row in arcpy.SearchCursor("tempTownRoad"):
                if row.RT_NUMBER not in town_route_num_dict[town]:
                    town_route_num_dict[town].append(row.RT_NUMBER)
                if row.STREET_NAM not in town_street_name_dict[town]:
                    town_street_name_dict[town].append(row.STREET_NAM)
        if route_number.find(" & ") > 0:
            arcpy.MakeFeatureLayer_management("tempTownRoad", "outlyr", "\"STREET_NAM\" IN ('" +
                                              "', '".join(i for i in route_number.split(" & ")) + "')")
            route_id_field = "STREET_NAM"
        else:
            if route_number in town_route_num_dict[town]:
                arcpy.MakeFeatureLayer_management("tempTownRoad", "outlyr", "(\"RT_NUMBER\" = '" + route_number +
                                                  "' OR \"ALTRTNUM1\" = '" + route_number + "' OR \"ALTRTNUM2\" = '" +
                                                  route_number + "' OR \"ALTRTNUM3\" = '" + route_number +
                                                  "' OR \"ALTRTNUM4\" = '" + route_number + "')")
                route_id_field = "RT_NUMBER"
            elif route_number in town_street_name_dict[town]:
                arcpy.MakeFeatureLayer_management("tempTownRoad", "outlyr", "\"STREET_NAM\" = '" +
                                                  route_number + "'")
                route_id_field = "STREET_NAM"

        if route_number.find(" & ") > 0:
            arcpy.MakeFeatureLayer_management("tempTownRoad", "outlyr", "\"STREET_NAM\" IN ('" +
                                              "', '".join(i for i in route_number.split(" & ")) + "')")
            route_id_field = "STREET_NAM"
        else:
            if route_number in town_route_num_dict[town]:
                arcpy.MakeFeatureLayer_management("tempTownRoad", "outlyr", "\"RT_NUMBER\" = '" + route_number + "'")
                route_id_field = "RT_NUMBER"
            elif route_number in town_street_name_dict[town]:
                arcpy.MakeFeatureLayer_management("tempTownRoad", "outlyr", "\"STREET_NAM\" = '" + route_number + "'")
                route_id_field = "STREET_NAM"
        arcpy.Dissolve_management("outlyr", temp_route, route_id_field)
        from_loc_lyr = determine_loc_lyr(from_location, town, output_geodatabase)
        end_loc_lyr = determine_loc_lyr(end_location, town, output_geodatabase)

        # without spatial analyst
        arcpy.Intersect_analysis([temp_route, from_loc_lyr], origin_point, "ONLY_FID", "1 Feet", "POINT")
        arcpy.Intersect_analysis([temp_route, end_loc_lyr], end_point, "ONLY_FID", "1 Feet", "POINT")

        # sort and find origin and end points
        arcpy.AddXY_management(origin_point)
        arcpy.AddXY_management(end_point)

        def sort_points_by_direction(pt, fc, pt_direction):
            point_type = os.path.basename(pt)
            coordinate_x = 0
            coordinate_y = 0
            if point_type == "origin":
                if pt_direction == 'N':
                    arcpy.Sort_management(pt, fc, [["POINT_Y", "ASCENDING"]])
                elif pt_direction == 'S':
                    arcpy.Sort_management(pt, fc, [["POINT_Y", "DESCENDING"]])
                elif pt_direction == 'W':
                    arcpy.Sort_management(pt, fc, [["POINT_X", "ASCENDING"]])
                elif pt_direction == 'E':
                    arcpy.Sort_management(pt, fc, [["POINT_X", "DESCENDING"]])
                elif pt_direction == "NE":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "LL")
                elif pt_direction == "NW":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "LR")
                elif pt_direction == "SE":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "UL")
                elif pt_direction == "SW":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "UR")
            elif point_type == "end":
                if pt_direction == 'N':
                    arcpy.Sort_management(pt, fc, [["POINT_Y", "DESCENDING"]])
                elif pt_direction == 'S':
                    arcpy.Sort_management(pt, fc, [["POINT_Y", "ASCENDING"]])
                elif pt_direction == 'W':
                    arcpy.Sort_management(pt, fc, [["POINT_X", "DESCENDING"]])
                elif pt_direction == 'E':
                    arcpy.Sort_management(pt, fc, [["POINT_X", "ASCENDING"]])
                elif pt_direction == "NE":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "UR")
                elif pt_direction == "NW":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "UL")
                elif pt_direction == "SE":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "LR")
                elif pt_direction == "SW":
                    arcpy.Sort_management(pt, fc, [["SHAPE", "ASCENDING"]], "LL")
            for rw in arcpy.SearchCursor(fc):
                if rw.OBJECTID == 1:
                    coordinate_x = rw.POINT_X
                    coordinate_y = rw.POINT_Y
            arcpy.Delete_management(fc)
            return coordinate_x, coordinate_y

        spatial_ref = arcpy.Describe(origin_point).spatialReference
        arcpy.CreateFeatureclass_management(output_geodatabase, "calibrate_points", "POINT", "", "", "", spatial_ref)
        arcpy.AddField_management(calibrate_points, "PointID", "SHORT")
        arcpy.AddField_management(calibrate_points, "RouteID", "TEXT")
        arcpy.AddField_management(calibrate_points, "Measure", "DOUBLE")
        rows = arcpy.InsertCursor(calibrate_points)

        # origin point
        origin_x, origin_y = sort_points_by_direction(origin_point, sort_fc, direction)
        new_row = rows.newRow()
        new_row.PointID = 1
        new_row.RouteID = route_number
        new_row.Measure = 0.0
        point = arcpy.Point(origin_x, origin_y)
        new_row.SHAPE = point
        rows.insertRow(new_row)

        # end point
        end_x, end_y = sort_points_by_direction(end_point, sort_fc, direction)
        new_row = rows.newRow()
        new_row.PointID = 2
        new_row.RouteID = route_number
        new_row.Measure = route_id_to_distance[(route, direction)] + from_distance + end_distance
        point = arcpy.Point(end_x, end_y)
        new_row.SHAPE = point
        rows.insertRow(new_row)

        arcpy.AddField_management(temp_route, "Length_Miles", "DOUBLE")
        arcpy.CalculateField_management(temp_route, "Length_Miles", "!SHAPE.LENGTH@MILES!", "PYTHON_9.3")
        arcpy.CreateRoutes_lr(temp_route, route_id_field, temp_routes, "ONE_FIELD", "Length_Miles")
        arcpy.CalibrateRoutes_lr(temp_routes, route_id_field, calibrate_points, "RouteID", "Measure", calibrated_route,
                                 "DISTANCE", "100 Feet")
        arcpy.MakeTableView_management(new_table, "tableView", "\"ROUTENUM\" = '" + route_number +
                                       "' AND \"TOWN\" = '" + town + "' AND \"FROM_LOC\" = '" + from_location +
                                       "' AND \"END_LOC\" = '" + end_location + "'")
        arcpy.MakeRouteEventLayer_lr(calibrated_route, route_id_field, "tableView", "ROUTENUM LINE START_MP END_MP", "routeEvents")
        arcpy.CopyFeatures_management("routeEvents", speed_route)
        speed_fc_list.append(speed_route)

        arcpy.Delete_management(calibrated_route)
        arcpy.Delete_management(calibrate_points)
        arcpy.Delete_management(end_point)
        arcpy.Delete_management(origin_point)
        arcpy.Delete_management(temp_route)
        arcpy.Delete_management(temp_routes)
        arcpy.Delete_management("outlyr")
        arcpy.Delete_management("routeEvents")
        arcpy.Delete_management("tableView")
        arcpy.Delete_management(from_loc_lyr)
        arcpy.Delete_management(end_loc_lyr)

    # output for municipal regs
    merged_fc = os.path.join(output_geodatabase, "merged")
    final_output = os.path.join(output_geodatabase, town.replace(" ", "_") + str(regulation))
    arcpy.Merge_management(speed_fc_list, final_output)

    # remove duplicates
    arcpy.Dissolve_management(merged_fc, final_output, ["REGULATION", "TOWN", "ROUTENUM", "DIRECTION", "FROM_LOC",
                                                        "FROM_DISTANCE_FEET", "END_LOC", "END_DISTANCE_FEET",
                                                        "DISTANCE", "SPEED", "AMENDMENT", "TIME_PERIOD", "LINK",
                                                        "START_MP", "END_MP", "Last_Segment", "JURISDICTION", "NOTES"])

    for s in speed_fc_list:
        arcpy.Delete_management(s)

    arcpy.Delete_management(new_table)
    arcpy.Delete_management(merged_fc)
    return

base_gdb = r"U:\Projects\Speed_Regulations\Base.gdb"
output_gdb = r"U:\Projects\Speed_Regulations\test\test.gdb"
road_inventory = r"U:\Projects\Speed_Regulations\updateRoadInventory.gdb\Roads"
jurisdiction_number = 2

# Create speed regulations
district_directory = r"U:\Projects\Speed_Regulations\test\DISTRICT 1"
for sub_directory, directory, files in os.walk(district_directory):
    for f in files:
        file_path = os.path.join(sub_directory, f)
        create_speed_regulation(file_path, base_gdb, output_gdb, road_inventory, jurisdiction_number)
