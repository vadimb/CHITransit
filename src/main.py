import transitfeed
import sys
from optparse import OptionParser
from xlrd import open_workbook
from schedulebuilder import ScheduleBuilder



def main():
    times_book = open_workbook('../data/schedule/excel/times.xls', formatting_info=True)
    schedules_book = open_workbook('../data/schedule/excel/schedules.xls', formatting_info=True)
    
    # args=__console();
    # input_path = args[0]
    input_path='../data/schedule/GTFS'
    # output_path = args[1]
    output_path='../data/schedule/gtfs_out'
    loader = transitfeed.Loader(input_path)
    schedule_in = loader.Load()
    loader = transitfeed.Loader(output_path)
    schedule_out = loader.Load()
    route=schedule_out.routes['1']
    route.trips[0].ClearStopTimes()
    

    builder = ScheduleBuilder(schedules_book, times_book,schedule_in,schedule_out)
    builder.build();
    schedule_out.Validate()
    schedule_out.WriteGoogleTransitFeed('../data/schedule/gtfs_out/out.zip')     
    
def __console():
    parser = OptionParser(
                          usage="usage: %prog [options] input_feed output_feed",
      version="%prog " + transitfeed.__version__)
    parser.add_option("-l", "--list_removed", dest="list_removed",
                    default=False,
                    action="store_true",
                    help="Print removed stops to stdout")
    (options, args) = parser.parse_args()
    if len(args) != 2:
        print >> sys.stderr, parser.format_help()
        print >> sys.stderr, "\n\nYou must provide input_feed and output_feed\n\n"
        sys.exit(2)
    return args

if __name__ == "__main__":
    main()
