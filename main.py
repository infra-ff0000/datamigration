__author__ = 'infraff0000'

# uncomment the next line when we are ready to batch process a directory!
import os

from xls_processor import XLSProcessor

# declare particulars
source_dir = 'src/disc-080514'
master_file = "src/AgileExample3_keep.xlsx"
results_saved_to = "src/better_all_results_aria.xls"

# declare variables
result_list = []

#  set up heading columns
result_list_headers = ['SKU', 'new status', 'site change', 'proposed web action', 'current web status']
result_list.append(result_list_headers)


# define helper functions
def returnProposedWebStatus(what_new_status):
    output = what_new_status + " (New Status)"
    if what_new_status == 'DIS- Discontinued':
        output = 'remove from web'

    if what_new_status == 'PND- Pending Discontinuation':
        output = 'pending removal'

    if what_new_status == 'ACT- Active':
        output = 'active on web'

    return output


def is_empty(any_structure):
    if any_structure:
        print('Not empty.')
        return False
    else:
        print('Is empty.')
        return True

current_source = XLSProcessor()
master = XLSProcessor()
# process directories
for subdir, dirs, files in os.walk(source_dir):
    for file in files:
        # create a new object for each file
        if file != ".DS_Store":
            current_file = os.path.join(subdir, file)
            print "Processing " + current_file

            # get source to work with, we will
            current_source.loadXLS(current_file)
            current_source.isolateTargetedRows(17)
            current_source.removeDuplicatesFromList()

            # instantiate master sku processor
            if is_empty(master.source_dictionary):
                master.loadXLS(master_file)
                master.compileMasterSourceIntoDictionary()

            #process items
            for item in current_source.source_dictionary:
                current_item = []
                current_sku = current_source.source_dictionary[item][1]
                print current_sku
                print "processing: " + str(current_sku)

                #insert the sku first
                current_item.insert(0, current_sku)

                # what is the new status?
                current_item.insert(1, current_source.source_dictionary[item][9])

                # the default web action is "everywhere"
                current_item.insert(2, 'everywhere')

                #based on the web status, what is the proposed action
                current_item.insert(3, returnProposedWebStatus(current_source.source_dictionary[item][9]))

                # find current item in master master list
                if item in master.source_dictionary:
                    print "we have a match: master source is "
                    print master.source_dictionary[item]
                    current_item.insert(4, master.source_dictionary[item][2])

                    # there might be another instance, so we need to
                    for key, value in master.source_dictionary.iteritems(): # iter on both keys and values
                        if key.startswith(current_sku):
                            current_item.insert(4, current_item[4] + "; " + key + " is listed as " + value[2])
                            print key, value
                else:
                    print "we don't have a match"
                    current_item.insert(4, 'N/A')

                print current_item
                print

                result_list.append(current_item)
            current_source.clearItems()

#save file
results = XLSProcessor()
results.saveResultlistAsXLS(result_list, results_saved_to)

print "Finished :) Check out " + results_saved_to
