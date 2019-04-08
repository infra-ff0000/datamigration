__author__ = 'infraff0000'

# uncomment the next line when we are ready to batch process a directory!
import os

from xls_processor import XLSProcessor

# declare particulars
source_dir = 'src/src_test/'
source_file = 'agile_MCO1037455.xls'
master_file = "src/AgileExample3_keep.xlsx"

# declare variables
# define where we are keeping the final results
result_list = []


# define helper functions
def returnProposedWebStatus(what_new_status):
    if what_new_status == 'DIS- Discontinued':
        return 'remove from web'

    if what_new_status == 'PND- Pending Discontinuation':
        return 'pending removal'

    if what_new_status == 'ACT- Active':
        return 'active on web'


# save a local copy
# master_dictionary = master.source_dictionary
# master.clearItems()

source = XLSProcessor()
# process directories

# for subdir, dirs, files in os.walk(source_dir):
#     for file in files:
#         # create a new object for each file
#         if file != ".DS_Store":
#current_file = os.path.join(subdir, file)
current_file = source_dir + source_file
print "Processing " + current_file

# get source to work with, we will
source.loadXLS(current_file)
source.isolateTargetedRows(17)
source.removeDuplicatesFromList()

# instantiate master sku processor
master = XLSProcessor()
master.loadXLS(master_file)
master.compileMasterSourceIntoDictionary()

#  set up heading columns
result_list_headers = ['SKU', 'new status', 'site change', 'proposed web action', 'current web status']
result_list.append(result_list_headers)
print result_list
# a helper function to determine proposed web status

#process items

for item in source.source_dictionary:
    current_item = []
    current_sku = source.source_dictionary[item][1]
    print current_sku
    print "processing: " + str(current_sku)

    #insert the sku first
    current_item.insert(0, current_sku)

    # what is the new status?
    current_item.insert(1, source.source_dictionary[item][9])

    # the default web action is "everywhere"
    current_item.insert(2, 'everywhere')

    #based on the web status, what is the proposed action
    current_item.insert(3, returnProposedWebStatus(source.source_dictionary[item][9]))

    # find current item in master master list
    if item in master.source_dictionary:
        print "we have a match: master source is "
        print master.source_dictionary[item]
        current_item.insert(4, master.source_dictionary[item][2])

        # there might be another instance, so we need to
        # for key, value in master.source_dictionary.iteritems(): # iter on both keys and values
        #     if key.startswith(current_sku):
        #         current_item.insert(4, current_item[4] + "; " + key + " is listed as " + value[2])
        #         print key, value
    else:
        print "we don't have a match"
        current_item.insert(4, 'N/A')

    print current_item
    print

    result_list.append(current_item)

print result_list

#save file
results = XLSProcessor()
results.saveResultlistAsXLS(result_list, "src/results_from_" + source_file)
print "results saved to : src/results_from_" + source_file
