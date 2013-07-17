import datetime
import transitfeed
class ScheduleBuilder():    
    SCHEDULE_CELLS=[[[workday for workday in xrange(1, 9)],[workday for workday in xrange(26, 34)]],[[ saturday for saturday in xrange(9, 16)],[ saturday for saturday in xrange(34, 41)]],[[ sunday for sunday in xrange(16, 23)],[ sunday for sunday in xrange(41, 47)]]] 
    HOUR_CELL = 0;  # A
    HOUR_ROW = 7;
    GRAY = "c0c0c0";
    STATION_ID_CELL = 0;  # A
    STATION_ID_ROW_OFFSET = 3;  # E
    
    def __init__(self, schedulesWorkbook, timesWorkbook,in_schedule,out_schedule):
        self.time_5am = datetime.datetime.now().replace(hour=5, minute=0, second=0, microsecond=0)
        self.time_2pm = datetime.datetime.now().replace(hour=14, minute=0, second=0, microsecond=0)
        self.time_7pm30 = datetime.datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        self.time_12pm = datetime.datetime.now().replace(hour=23, minute=59, second=0, microsecond=0)              
        self.time_9am = datetime.datetime.now().replace(hour=9, minute=0, second=0, microsecond=0)
        self.time_8pm = datetime.datetime.now().replace(hour=20, minute=0, second=0, microsecond=0)
        self.time_0am = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        self.time_0am=datetime.timedelta(hours=36)
        self.in_schedule=in_schedule
        self.out_schedule=out_schedule
        self.scheduleWorkbook = schedulesWorkbook
        self.timeWorkbook = timesWorkbook  
        
    def build(self):  
        for sheet in self.scheduleWorkbook.sheets():
            route=self.in_schedule.routes[sheet.name]                     
            trips=route.trips
            
            firstTrip=trips[0]
            #template stop time
            stop_times=firstTrip.GetStopTimes()
            self.__build_schedule(sheet, sheet.name,trips[:3],stop_times,0)
            
            firstTrip=trips[3]
            #template stop time
            stop_times=firstTrip.GetStopTimes()
            self.__build_schedule(sheet, sheet.name,trips[3:],stop_times,1) 
            

    def __build_schedule(self, sheet, route,trips,stop_times,direction):
                  
        #process 3 trips(workdays,saturday,sunday)
        for trip in trips:
            if direction:
                cell_index = 25
            else :                
                cell_index = ScheduleBuilder.HOUR_CELL
            
            row_index = ScheduleBuilder.HOUR_ROW
            hour = 0 
            # process day
            while (hour <= 23 and row_index<sheet.nrows and cell_index<sheet.ncols) :
                cell = sheet.cell(row_index, cell_index)
                
                if cell.value:
                    hour = int(cell.value)                    
                

                color = self.__cell_bg(self.scheduleWorkbook, sheet, row_index, cell_index)      
                # process hour
                while (color == self.__cell_bg(self.scheduleWorkbook, sheet, row_index, cell_index)):
                    # process minute
                    minutes = ''
                    for day in ScheduleBuilder.SCHEDULE_CELLS[int(trip.service_id)-1][int(direction)] :
                        cell_index = day
                        cell = sheet.cell(row_index, cell_index);
                        minutes = cell.value
                        if not minutes:
                            break;
                        route_out=self.out_schedule.routes[route] 
                        new_trip=route_out.AddTrip(headsign=trip.trip_headsign, service_period=trip.service_period)
                        workday=int(trip.service_id)==1             
                        self.__linkSchedule(int(hour), int(minutes), int(route),new_trip,stop_times,route_out,workday,direction) 
                        #route.AddTrip               
                    row_index += 1
                    cell_index = ScheduleBuilder.HOUR_CELL
                                        
                    if not minutes:
                        break;     

    def __linkSchedule(self, hour, minute, route,trip,stop_times,route1,workday=True,direction=0) :      
        timeSheet = self.timeWorkbook.sheet_by_name(str(route))
        if direction==0:
            first_row= ScheduleBuilder.STATION_ID_ROW_OFFSET+1
            last_row= int(timeSheet.cell(timeSheet.nrows-1, 0).value)-1
        else:
            first_row= int(timeSheet.cell(timeSheet.nrows-1, 0).value)
            last_row= int(timeSheet.cell(timeSheet.nrows-1, 1).value)-1
        
        if last_row-first_row+1!=len(stop_times):
            raise Exception("Stop times don't equal "+str(last_row-first_row+1)+"!="+str(len(stop_times))+"on route: "+str(route) +" direction:"+str(direction))

        bucket = self.__get_time_col(hour, minute, workday)
        time = datetime.datetime.now().replace(hour=hour, minute=minute, second=0, microsecond=0)
        begin_time=datetime.datetime.now().replace(hour=hour, minute=minute, second=0, microsecond=0)
        for station in range(first_row, last_row+1):       
            minutes = timeSheet.cell_value(station, bucket)  
            time = time + datetime.timedelta(minutes=minutes)
            t=str(time.time())            
            if(begin_time.hour>time.hour):
                time=datetime.timedelta(hours=time.hour,minutes=time.minute,seconds=time.second)                
                m, s = divmod(time.seconds+86400, 60)
                h, m = divmod(m, 60)
                t="%d:%02d:%02d" % (h, m, s)

            stop_time=stop_times[(station-first_row)]
            stop=transitfeed.StopTime(problems=None,stop=stop_time.stop,arrival_time=t, departure_time=t,stop_sequence=stop_time.GetFieldValuesTuple(trip.trip_id)[4])
            trip.AddStopTimeObject(stop)

    def __get_time_col(self, hour, minute, workday=True) :
        time = datetime.datetime.now().replace(hour=hour, minute=minute, second=0, microsecond=0)

        if time >= self.time_5am and time < self.time_2pm and workday:
            cell = 9        

        if time >= self.time_2pm and time < self.time_7pm30 and workday:
            cell = 10        

        if time >= self.time_7pm30 and time <= self.time_12pm and workday:
            cell = 11
            
        if time >= self.time_5am and time < self.time_9am and not workday:
            cell = 12        

        if time >= self.time_9am and time < self.time_8pm and not workday:
            cell = 13        

        if time >= self.time_8pm and time <= self.time_12pm and not workday:
            cell = 14        
        return cell
    
    def __cell_bg(self, workbook, sheet, row, col):
        xf_index = sheet.cell_xf_index(row, col)
        xf = workbook.xf_list[xf_index]
        return xf.background.pattern_colour_index