import os
import fnmatch
import win32com.client

def get_file_metadata(imgs_path, filesnames):
    # Path shouldn't end with backslash, i.e. "E:\Images\Paris"
    # filename must include extension, i.e. "PID manual.pdf"
    # Returns dictionary containing all file metadata.
	sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
	ns = sh.NameSpace(imgs_path)
	
	#metadata = ['Name', 'Size', 'Item type', 'Date modified', 'Date created']
	file_metadata_list = list()
	for i in filesnames: #create a list of dictionary 
		file_metadata = dict()
		item = ns.ParseName(str(i))
		attr_value = ns.GetDetailsOf(item, 3)# index location of 'Data Modifited" - 4 -> 3
		if attr_value:
			file_metadata[i] = attr_value
		file_metadata_list.append(file_metadata)
	
	return file_metadata_list


def sort_by_date(Meta_list):
	
	for index_1 in range(len(Meta_list)):
		for index_2 in range(len(Meta_list)):
			try:	
					#extract only the "Date Modified" values
					#split the values into 'Date', 'Month', 'Year'
				Meta_index_1 = list(Meta_list[index_1].values())[0].split('/')
				Meta_index_2 = list(Meta_list[index_2].values())[0].split('/')
				
				if(int(Meta_index_1[2][:4]) < int(Meta_index_2[2][:4])): #comparing Years
					temp = Meta_list[index_1]
					Meta_list[index_1] = Meta_list[index_2]
					Meta_list[index_2] = temp
				elif(int(Meta_index_1[2][:4]) == int(Meta_index_2[2][:4])):
					if(int(Meta_index_1[0][:4]) < int(Meta_index_2[0][:4])):#comparing Months
						temp = Meta_list[index_1]
						Meta_list[index_1] = Meta_list[index_2]
						Meta_list[index_2] = temp
					elif(int(Meta_index_1[0][:4]) == int(Meta_index_2[0][:4])):
						if(int(Meta_index_1[1][:4]) < int(Meta_index_2[1][:4])):#comparing Dates
							temp = Meta_list[index_1]
							Meta_list[index_1] = Meta_list[index_2]
							Meta_list[index_2] = temp
							
							
			except IndexError:
				print(f"The dictionary value is not as expected, for example: {Meta_list[index_1]}")
				return "Error Message function: 'sort_by_date'"
	return Meta_list
		
	
def main():
		#read files in dirctory
	while(1):
		dict_images = list()
		folder_name = input("Write pictures folder name with main.py location:")
		program_location = os.path.dirname(os.path.abspath(__file__))
		pictures_path = f"{program_location}\\{folder_name}"
		if((input(f"Is '{pictures_path}' the right path? 'y', 'n':	")).lower()=='n'):
			continue
		
		try:
			files_list = os.listdir(pictures_path)
			dict_images = get_file_metadata(pictures_path, files_list)
			sorted_list = sort_by_date(dict_images)
			
			if isinstance(sorted_list,str):# checking variable type matches 'str'
				print(sorted_list)
				return
			try:
				#new_name = input("header word/ letter for pictures: ")
				for i in range(len(sorted_list)):
				
					os.rename(f"{pictures_path}\\{list(sorted_list[i].keys())[0]}",  f"{pictures_path}\\{i+1}_{list(sorted_list[i].keys())[0]}")
			except FileExistsError:
				print("Couldn't overed existing file, try a different name")
		except FileNotFoundError:
			print("Folder not found")
			
		break #stop the while
			
	
if __name__ == "__main__": 
    main()