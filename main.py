import os
import fnmatch
import win32com.client

def get_file_metadata(imgs_path, filesnames):
    # Path shouldn't end with backslash, i.e. "E:\Images\Paris"
    # filename must include extension, i.e. "PID manual.pdf"
    # Returns dictionary containing all file metadata.
	sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
	print(imgs_path)
	ns = sh.NameSpace(imgs_path)
	
	#metadata = ['Name', 'Size', 'Item type', 'Date modified', 'Date created']
	file_metadata_list = list()
	for i in filesnames: #create a list of dictionary 
		file_metadata = dict()
		item = ns.ParseName(str(i))
		attr_value = ns.GetDetailsOf(item, 4)
		if attr_value:
			file_metadata[i] = attr_value
		#	print(file_metadata)
		file_metadata_list.append(file_metadata)
	return file_metadata_list

def sort_by_date(Meta_list):
	
	for index_1 in range(len(Meta_list)):
		for index_2 in range(len(Meta_list)):
			try:
				if(str(list(Meta_list[index_1].values())).split('/')[2][:4] < str(list(Meta_list[index_2].values())).split('/')[2][:4]):
					temp = Meta_list[index_1]
					Meta_list[index_1] = Meta_list[index_2]
					Meta_list[index_2] = temp
			except IndexError:
				print(f"The dictionary value is not as expected, for example: {Meta_list[index_1]}")
				return "Error Message function: 'sort_by_date'"
	return Meta_list
		
	
def main():
		#read files in dirctory
	dict_images = list()
	folder_name = input("Write pictures folder name with main.py location:")
	program_location = os.path.dirname(os.path.abspath(__file__))
	pictures_path = program_location+ f'\\{folder_name}'
	try:
		files_list = os.listdir(pictures_path)
		#filter files that dont match, anything but *.jpeg or *.jpg
		#filtered_list = [i for i in files_list if  (fnmatch.fnmatch(i,'*.jpg') or fnmatch.fnmatch(i,'*.jpeg'))]
		dict_images = get_file_metadata(pictures_path, files_list)
		sorted_list = sort_by_date(dict_images)
		if isinstance(sorted_list,str):# checking variable type matches 'str'
			print(sorted_list)
			return
		try:
			new_name = input("header word/ letter for pictures: ")
			for i in range(len(sorted_list)):
				os.rename(pictures_path + '\\'+ str(list(sorted_list[i].keys()))[2:-2],pictures_path +  f"\\{new_name}{i+1}.jpg")
		except FileExistsError:
			print("Couldn't overed existing file, try a different name")
	except FileNotFoundError:
		print("Folder not found")
		
	
	
	
if __name__ == "__main__": 
    main()