# https://projects.propublica.org/nonprofits/api/v2/search.json?q=propublica


from bs4 import BeautifulSoup
from bs4 import SoupStrainer

import xmltodict
import xlsxwriter 

import xlrd

# import openpyxl module 
import openpyxl
from openpyxl import load_workbook

import requests
import re

#ein = input("EIN for nonprofit:")

'''
eins = [ "010372148", "133978518", "381405282", "450374475", "470665946", "660235625", "920037099", "941719894", "942690415", "980078729", "920037099"]

eins_debug = ["010372148", "133978518", "920037099"]
revenue_debug = [ 222546799 ,  23803439 , 50835652]

tot_rev_cy_2009 = [ 209048486 ,  19326854 ,  8927540 ,  555733 ,  21102294 ,  5907063 ,  34615095 ,  145540880 ,  77166297 ,  51662705 ]
tot_rev_cy_2010 = [ 216648084 ,  23131766 ,  6540403 ,  423308 ,  21578804 ,  6455871 ,  38171806 ,  148235603 ,  93545410 ,  54149051 ]
tot_rev_cy_2011 = [ 217882913 ,  22186860 ,  13702248 ,  488219 ,  22690530 ,  8835935 ,  43053727 ,  140559402 ,  102963440 ,  5768688 ]
tot_rev_cy_2012 = [ 222546799 ,  23803439 ,  11181573 ,  598286 ,  23560976 ,  10176570 ,  50835652 ,  149419284 ,  113373196 ,  57853984 , 50835652]
tot_rev_cy_2013 = [ 237325771 ,  28525784 ,  12846129 ,  638982 ,  28898918 ,  5183126 ,  54972418 ,  157183976 ,  132777853 ,  59810651 ]
tot_rev_cy_2014 = [ 251737826 ,  33283168 ,  14149530 ,  473444 ,  31408452 ,  5669521 ,  61294402 ,  162944541 ,  146947928 ,  61187887 ]
tot_rev_cy_2015 = [ 278270243 ,  40420032 ,  12984899 ,  217166 ,  28019184 ,  4509337 ,  66526394 ,  177073446 ,  165241230 ,  63615221 ]
'''

all_125_eins = ['010215911', '010276272', '010284446', '010370309', '010372148', '010378758', '030179404', '042024022', '042103542', '042103550', '042104144', '042125005', '042197449', '042429311', '042814018', '043377720', '046074663', '046113618', '050351121', '060646702', '060646960', '060653119', '060677159', '060733693', '061051768', '061423598', '066002281', '112626155', '131623850', '131624048', '131635293', '131878953', '131974191', '133004747', '133913393', '133978518', '134043587', '136165593', '141501404', '141715960', '146022764', '160759798', '161025787', '161043694', '200097189', '200253310', '221487252', '221487307', '221720078', '222083588', '222367416', '223331880', '231352354', '231496890', '232048152', '232220051', '232657495', '232703530', '236437695', '237062502', '237066616', '237099149', '237161473', '237182582', '237267810', '237314364', '237373091', '237398279', '237410494', '237413415', '240772407', '250585280', '250979346', '250986051', '251340370', '251371578', '300406260', '311333812', '311551871', '311680916', '311688648', '330061828', '330828761', '330898712', '340714426', '340714670', '341314846', '341372076', '341522484', '341611055', '341839195', '341908579', '346007227', '350985964', '351362157', '351898055', '361892375', '362170786', '363535053', '366008929', '366108515', '370673513', '371079626', '371179056', '381359513', '381398840', '381405282', '381459365', '382394044', '382476777', '382709547', '386094484', '390806395', '391032051', '391157876', '391691578', '410695462', '410874507', '410946789', '411360294', '411390349', '411854826', '420680323', '420680372', '420843389']


tot_rev_cy_2009 = ['110673680', '642968', '69928031', '4568988', '209048486', '12386189', '14065438', '1774178', '244624704', '73377350', '17838185', '219758', '100118978', '23466558', '10255308', '427569', '-39593', '368908', '8470386', '14543484', '25684402', '45824916', '9976261', '26232445', '569700', '523195', '3390363', '227517920', '6744973', '14621910', '49185151', '13395168', '616181002', '72382451', '4042072', '19326854', '9733455', '1012853', '3763555', '894966', '518872', '52615059', '4775033', '874993', '910995', '2963983', '20134383', '587190505', '3090218', '3074766', '2100610', '11375792', '68870284', '14324982', '825270', '16503559', '4222351', '1140249', '50101671', '1349282', '6022519', '19748326', '6235132', '31284941', '3867222', '26347843', '15100804', '1274554', '1542942', '368978', '216424449', '71560636', '70183022', '25680247', '45097649', '420876', '1701693', '756807', '23396743', '1691967', '1459901', '28726165', '6699015', '17211265', '26753945', '52210140', '918815', '426286', '3365102', '10973997', '29206884', '-5821733', '2610078', '143685987', '13159319', '4825866', '895625185', '34321066', '2185376', '63483356', '10051341', '69320751', '32431836', '13836235', '44397618', '1165272', '8927540', '28830920', '14802368', '-254741', '587629', '566733', '122957032', '1164540', '93012771', '5266542', '14464016', '9438337', '9115308', '2230912', '4993260', '349930', '60984987', '64203531', '15248088']

tot_rev_cy_2010 = ['122900842', '509721', '72414792', '4496356', '216648084', '12994317', '16631751', '2387504', '280327817', '80659639', '20294317', '252705', '104697677', '25543661', '6473357', '452419', '555469', '502177', '8209633', '13880130', '23186005', '42968171', '9757334', '27613779', '534132', '521701', '2505898', '198407862', '10105487', '15362694', '49780019', '14141934', '627468115', '42538653', '4020269', '23131766', '9635517', '1016420', '3791222', '729935', '609121', '56731755', '5059682', '977840', '1934659', '4761721', '22030891', '622284006', '3037745', '3394323', '2087440', '11748529', '74151095', '12404914', '1776686', '9362560', '4708362', '1027474', '47460681', '1308656', '5666237', '25098624', '6496724', '36307108', '4866818', '28741854', '12434193', '1531969', '1704309', '592631', '233003584', '77455262', '67450652', '27295970', '47644475', '94914', '1577050', '565971', '24518142', '3131031', '1510696', '28803227', '11391916', '14667240', '30623642', '61405889', '734365', '1258198', '3371729', '2023911', '28497239', '20700819', '2463777', '151620205', '11042289', '7404649', '1046518213', '37520308', '2445761', '55903106', '11987142', '71635606', '33470296', '16026179', '58958407', '995715', '6540403', '28293560', '17035368', '723109', '571762', '578043', '133572374', '1337397', '100509661', '7017272', '21139960', '8644405', '9102678', '2628805', '5149542', '334858', '59171945', '81334145', '16074646']

tot_rev_cy_2011 = ['126160636', '430058', '73602550', '5806251', '217882913', '14750944', '16465207', '1798227', '325110846', '97219404', '19190201', '330027', '107920617', '26108814', '3766255', '487517', '847541', '287176', '6985866', '13875156', '40700270', '36316665', '9228002', '27237761', '521451', '546703', '3016157', '206522629', '7057143', '14466797', '50778075', '16208495', '579258360', '59724538', '4300830', '22186860', '14545779', '959183', '3410088', '729657', '610308', '62148128', '5777984', '934105', '3954108', '5198829', '21245593', '624131028', '2832096', '3428107', '2726075', '10649725', '75131609', '13196883', '1687331', '17257138', '4339705', '1342638', '46059756', '1336137', '7609297', '23025607', '6477562', '37522603', '9381474', '34450300', '13085470', '1463250', '1718959', '1065174', '248338908', '80772707', '64184272', '24600962', '48343691', '94665', '1587371', '1394071', '25137927', '4066368', '1522238', '28677188', '11052908', '43148413', '23807244', '50346723', '647721', '1985384', '3642294', '4413125', '31750872', '52628401', '2463176', '164483382', '10865845', '8556775', '2077027269', '40024514', '2576804', '54279633', '6230236', '73639199', '34008361', '13206310', '33364638', '1150140', '13702248', '28510311', '33482440', '889694', '425347', '615685', '134171848', '1395670', '105529940', '3555917', '16753165', '7591361', '9248838', '4072487', '4986978', '273951', '54573305', '63688131', '17302414']

tot_rev_cy_2012 = ['136494576', '323359', '78963358', '5979592', '222546799', '14613790', '14925860', '2447130', '356118724', '98057222', '18299522', '296633', '99763852', '27752895', '2838718', '424492', '870339', '1469830', '6607348', '14941259', '24368847', '43301645', '10918403', '33997334', '456147', '554103', '1978330', '183189420', '6855584', '18249706', '52398185', '17369390', '606996477', '40285288', '4888724', '23803439', '10878204', '1042270', '3207779', '362320', '812745', '66627739', '6698961', '1094360', '9046361', '9398261', '22864953', '661418701', '2617587', '3768747', '3190829', '10462684', '77598651', '14525244', '1417500', '15640045', '4564719', '1281124', '45289276', '1341027', '8912295', '26140458', '6490846', '34137504', '8324251', '37910321', '13360020', '1479065', '1629208', '984431', '251153748', '77589321', '52707684', '25182320', '47569636', '66520', '1551170', '2397619', '24575773', '3199640', '1593978', '27359517', '10983193', '10219009', '31718418', '56529823', '748710', '1765098', '3429412', '15078111', '35343319', '85628334', '3104402', '169929334', '11730988', '9411059', '17210587', '42249472', '2805747', '60743107', '5496331', '85681509', '33415907', '28476776', '57340312', '812938', '11181573', '28969843', '20017249', '1434125', '499936', '635699', '129493990', '1807946', '103952476', '5817399', '16178184', '6888997', '9102424', '3000917', '5072534', '324226', '77156234', '62103019', '17480151']

tot_rev_cy_2013 = ['147444854', '451200', '81089005', '6418238', '237325771', '19121874', '16312508', '4869429', '375381465', '100653882', '18655503', '1081152', '99673988', '29450780', '2661712', '550451', '1256538', '222580', '6767452', '15250685', '24579996', '50947712', '10762418', '32252008', '516077', '698179', '3354525', '174759706', '9975029', '15519995', '55091496', '17735493', '631049304', '38943816', '4528753', '28525784', '12423050', '1169553', '3196255', '720815', '691137', '71104111', '7605911', '1234349', '13493744', '10139454', '23388794', '663253034', '3004249', '3913143', '3424137', '11074499', '88033092', '9872533', '1396285', '20542094', '4761743', '1430663', '40902914', '1309481', '10206669', '25649926', '6691459', '32966905', '4213586', '37175065', '14126546', '2198664', '1468902', '1192544', '256785507', '75854556', '42992016', '24316330', '47666978', '89434', '1648036', '806925', '25307775', '5064258', '1586902', '35233217', '10941332', '21494184', '28189567', '58869414', '896429', '1422234', '3554689', '5710945', '25244883', '18190308', '3799316', '191109000', '11774229', '9365511', '16907419', '44376054', '2652865', '90532766', '5380540', '93552771', '33084921', '16594465', '53023137', '855975', '12846129', '33491747', '23665325', '1066432', '512148', '666936', '136936736', '2207504', '111290605', '7130394', '18540059', '7223438', '9454487', '4202002', '5243276', '325447', '62528992', '68572391', '16940476']

tot_rev_cy_2014 = ['166911061', '398922', '88764244', '7158666', '251737826', '18331815', '17728388', '3231095', '396785932', '106807329', '18884939', '1525612', '106520889', '30786399', '3526534', '533306', '1082695', '3373339', '7559630', '14251147', '34522108', '47629718', '10974019', '36072710', '649309', '757947', '2487722', '184701438', '8974106', '17225488', '55785531', '20947318', '680259090', '37422921', '4273737', '33283168', '31533905', '1128488', '3350260', '325966', '837474', '74661376', '7463200', '1533905', '19309416', '10304545', '30195083', '677137484', '2989974', '3747455', '3428871', '11775140', '86106352', '9808885', '1543375', '18696371', '4564064', '1523355', '41887235', '1375292', '12844786', '26683170', '6967965', '20096038', '4570161', '52959255', '14180617', '2965430', '1546531', '946559', '267244714', '66051909', '44028132', '24045418', '49242297', '119728', '1601085', '876618', '23824009', '3281018', '1587722', '27536347', '10878819', '8205670', '28666573', '51591946', '1005051', '1150167', '3524038', '3026583', '23164372', '9456826', '3233251', '198061057', '11763353', '8794508', '18196678', '46880836', '2847199', '85425939', '3961302', '92116143', '32714351', '15165740', '79243633', '1104832', '14149530', '31253257', '27094971', '971446', '490775', '662340', '132045126', '2255647', '106161399', '8186226', '17813558', '8549727', '11244876', '3758670', '5328664', '356098', '73185269', '76273426', '18075230']

tot_rev_cy_2015 = ['190659711', '321793', '97503169', '7470291', '278270243', '19390193', '15656308', '3181827', '329573554', '84738790', '18995390', '1982519', '110973930', '31663105', '3508917', '346235', '719056', '3164141', '7672583', '13677944', '44606874', '46281851', '11096749', '36200199', '710833', '1545348', '3209349', '203162835', '9086793', '17818609', '55785292', '22639386', '701001252', '60014592', '4762667', '40420032', '19709765', '1268979', '3546379', '609393', '808891', '74285791', '7571405', '1602763', '19458887', '7073443', '28875690', '672424783', '3419783', '5099604', '3680880', '11972835', '91212398', '10258309', '1325839', '23817632', '3552989', '1385120', '41923269', '1398498', '8578456', '29055481', '7143820', '36611563', '3427861', '50892509', '13537620', '3090597', '1480344', '6381704', '261750893', '62913776', '43012266', '26143695', '48445675', '112269', '1800007', '1520342', '28003497', '-165984', '1580230', '26041099', '11602677', '9848663', '27705856', '47136490', '1108438', '1226190', '3864110', '10920069', '20134510', '10437949', '3342383', '209945833', '11972381', '8020267', '22649705', '50205024', '3104485', '70225003', '7439476', '85486520', '32673970', '14233837', '61610348', '1038847', '12984899', '35765475', '21853681', '862725', '409841', '762345', '134230652', '2105906', '106736768', '11330311', '18430551', '9108002', '11115393', '3917964', '6198810', '392717', '83728369', '73505221', '19394508']





#first_batch_of_fifty_2009_eins = all_125_eins[0:50]
first_batch_of_fifty_2009_eins = all_125_eins[0:2]
first_bath_of_fifty_2009_revenues = tot_rev_cy_2009[0:2]

next_batch_of_fifty_2009_eins = all_125_eins[50:101]
next_bath_of_fifty_2009_revenues = tot_rev_cy_2009[50:101]
'''
print(len(all_125_eins))
print(len(tot_rev_cy_2009))

print(len(first_batch_of_fifty_2009_eins))
print(first_batch_of_fifty_2009_eins[0])
print(first_batch_of_fifty_2009_eins[49])

print(len(first_bath_of_fifty_2009_revenues))
print(first_bath_of_fifty_2009_revenues[0])
print(first_bath_of_fifty_2009_revenues[49])
'''


output_path = "/Users/ssd/Downloads/template_round4.xlsx"
book = openpyxl.load_workbook(output_path)
sheet = book["year12"]

#finding which letter column each header is located in...

ein_column_letter = ""
total_revenue_column_letter = ""
name_column_letter = ""
position_column_letter = ""
salary_column_letter = ""
extra_comp_column_letter = ""
other_comp_column_letter = ""
hours_per_week_column_letter = ""
trustee_or_director_column_letter = ""
institutional_trustee_column_letter = ""
officer_column_letter = ""
key_employee_column_letter = ""
highest_paid_column_letter = ""
former_column_letter = ""
tags_column_letter = ""
length_column_letter = ""
person_id_column_letter = ""



letter_list = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
i = 1
for letter in letter_list:

	if sheet[letter + str(i)].value =="EIN":
		ein_column_letter = letter  
	if sheet[letter + str(i)].value =="total_revenue_cy":
		total_revenue_column_letter = letter
	if sheet[letter + str(i)].value =="name":
		name_column_letter = letter
	if sheet[letter + str(i)].value =="position":
		position_column_letter = letter
	if sheet[letter + str(i)].value =="salary":
		salary_column_letter = letter
	if sheet[letter + str(i)].value =="extracomp":
		extra_comp_column_letter = letter
	if sheet[letter + str(i)].value =="othercomp":
		other_comp_column_letter = letter
	if sheet[letter + str(i)].value =="hoursperweek":
		hours_per_week_column_letter = letter
	if sheet[letter + str(i)].value =="trusteedirector":
		trustee_or_director_column_letter = letter
	if sheet[letter + str(i)].value =="institutionaltrustee":
		institutional_trustee_column_letter = letter
	if sheet[letter + str(i)].value =="officer":
		officer_column_letter = letter
	if sheet[letter + str(i)].value =="keyemployee":
		key_employee_column_letter = letter
	if sheet[letter + str(i)].value =="highestpaid":
		highest_paid_column_letter = letter
	if sheet[letter + str(i)].value =="former":
		former_column_letter = letter
	if sheet[letter + str(i)].value == "tags":
		tags_column_letter = letter
	if sheet[letter + str(i)].value == "length":
		length_column_letter = letter
	if sheet[letter + str(i)].value == "person_id":
		person_id_column_letter = letter

#print(ein_column_letter)

i = 0
counter = 0
position = ""

reportable_comp_from_organization = 0
reported_comp_from_related = 0
other_comp = 0
avg_hrs_per_week = 0
trustee_or_director = ""
officer = ""
highest_compensated = ""
key_employee = ""
formerly_employed = ""


for ein in first_batch_of_fifty_2009_eins:

	print("EIN NUMBER:  " + str(i + 1))


	total_revenue = first_bath_of_fifty_2009_revenues[i]
	r = requests.get("https://projects.propublica.org/nonprofits/organizations/" + str(ein))
	soup = BeautifulSoup(r.text, 'html.parser')

	two_filings_found_for_year = False 

	single_filings = soup.find_all('div', {'class' :'single-filing cf'})
	xml_url = []
	for single in single_filings:
		a = single.find('a', {'class' :'action xml'})
		if a != None:
			xml_url.append((a['href']))
		subject_options = [i.findAll('option') for i in single.findAll('select', attrs = {'class': 'action xml'} )]
	

		#print(subject_options)
		if subject_options != []:
			print("found 2 filings occurence ")
			#import pdb; pdb.set_trace()
			soup2 = BeautifulSoup(str(subject_options[0][1]), 'html.parser')
			link2 = soup2.find("option")['data-href']
			soup3 = BeautifulSoup(str(subject_options[0][2]), 'html.parser')
			link3 = soup3.find("option")['data-href']

			load_first_xml = requests.get(link2)
			load_second_xml = requests.get(link3)

			print(link2)
			print(link3)

			first_xml_doc = BeautifulSoup(load_first_xml.text, 'html.parser')
			second_xml_doc = BeautifulSoup(load_second_xml.text, 'html.parser')


			tag1 = first_xml_doc.find(re.compile("cytotalrevenueamt"))
			tag2 = first_xml_doc.find(re.compile("totalrevenuecurrentyear"))
			tag3 = second_xml_doc.find(re.compile("cytotalrevenueamt"))
			tag4 = second_xml_doc.find(re.compile("totalrevenuecurrentyear"))


			tag1_true = False
			tag2_true = False
			tag3_true = False
			tag4_true = False

			if tag1 != None:
				if tag1.text == str(total_revenue):
					tag1_true = True
			if tag2 != None:
				if tag2.text == str(total_revenue):
					tag2_true = True
			if tag3 != None:
				if tag3.text == str(total_revenue):
					tag3_true = True
			if tag4 != None:
				if tag4.text == str(total_revenue):
					tag4_true = True

			first_xml_is_valid = (tag1_true) or (tag2_true)
			second_xml_is_valid = (tag3_true) or (tag4_true)

			if first_xml_is_valid & second_xml_is_valid:
				two_filings_found_for_year = True

				ein_column_value = ein_column_letter + str(counter+2)
				sheet[ein_column_value] = ein
				book.save(output_path)

				revenue_column_value = total_revenue_column_letter + str(counter+2)
				sheet[revenue_column_value] = total_revenue
				book.save(output_path)

				name_column_value = name_column_letter + str(counter+2)
				sheet[name_column_value] = "2 FILINGS for this year"
				book.save(output_path)

				counter +=1

				break

			elif first_xml_is_valid & (not second_xml_is_valid):
				xml_url.append(link2)

			elif not first_xml_is_valid & second_xml_is_valid:
				xml_url.append(link3)






	target_xml = ""
	for each_xml in xml_url:
		x = requests.get(each_xml)
		soup = BeautifulSoup(x.text, 'html.parser')
		y = soup.find(re.compile("cytotalrevenueamt"))
		z = soup.find(re.compile("totalrevenuecurrentyear"))


		if y!= None:
			if str(total_revenue) == y.text:
				target_xml = each_xml
				print(y.text)
				print("got the right xml: " + each_xml)
				break
		if z!= None:
			if str(total_revenue) == z.text:
				target_xml = each_xml
				print(z.text)
				print("found the right xml: " + each_xml)
				break
	
	if target_xml != "":
		correct_file = requests.get(target_xml)
		soup = BeautifulSoup(correct_file.text, 'html.parser')
		board_directors_names = soup.find_all(re.compile("form990partviisectiona"))


		#check to see if there is at least one occurrence of the attribute

		is_trustee_or_director = False
		is_officer = False
		is_highest_paid = False
		is_reported_compensated = False
		is_related_comp = False
		is_other = False
		is_avg_hours = False
		is_former = False
		is_key = False

		
		for each_name_block in board_directors_names:

			if each_name_block.find(re.compile("individualtrusteeordirector")) != None or each_name_block.find(re.compile("individualtrusteeordirectorind")) != None:
				is_trustee_or_director = True
			if each_name_block.find(re.compile("officer")) != None:
				is_officer = True
			if each_name_block.find(re.compile("highestcompensatedemployee")) != None:
				is_highest_paid = True
			if each_name_block.find(re.compile("reportablecompfromorganization")) != None or each_name_block.find(re.compile("reportablecompfromorgamt")) != None:
				is_reported_compensated = True
			if each_name_block.find(re.compile("reportablecompfromrelatedorgs")) != None or each_name_block.find(re.compile("reportablecompfromrltdorgamt")) != None:
				is_related_comp = True
			if each_name_block.find(re.compile("averagehoursperweek")) != None:
				is_avg_hours = True
			if each_name_block.find(re.compile("othercompensation")) != None:
				is_other = True
			if each_name_block.find(re.compile("formerofcremployeeslistedind")) != None:
				is_former = True
			if each_name_block.find(re.compile("keyemployeeind")) != None or each_name_block.find(re.compile("keyemployee")) != None:
				is_key = True

		


		

		count_people = 1


		for each_name_block in board_directors_names:
			#print(each_name_block)

			tag_list = []
			# print(ein + "_" + str(total_revenue) + "    PERSON")
			for child in each_name_block:
				child_string = str(child)
				tag = child_string[child_string.find("<") + 1:child_string.find(">")]
				if tag != "":
					tag_list.append(tag)
			tags_column_value = tags_column_letter + str(counter+2)
			sheet[tags_column_value] = repr(tag_list)
			length_column_value = length_column_letter + str(counter+2)
			sheet[length_column_value] = len(tag_list)
			book.save(output_path)


			person_id_column_value = person_id_column_letter + str(counter+2)
			sheet[person_id_column_value] = count_people
			book.save(output_path)


			ein_column_value = ein_column_letter + str(counter+2)
			sheet[ein_column_value] = ein
			book.save(output_path)


			revenue_column_value = total_revenue_column_letter + str(counter+2)
			sheet[revenue_column_value] = total_revenue
			book.save(output_path)


			if each_name_block.find(re.compile("nameperson")) != None:
				person_name = each_name_block.find(re.compile("nameperson")).text
			elif each_name_block.find(re.compile("personnm")) != None:
				person_name = each_name_block.find(re.compile("personnm")).text
			elif each_name_block.find(re.compile("personnm")) == None and each_name_block.find(re.compile("nameperson")) == None:
				person_name = "no name found with current tags in code"
			name_column_value = name_column_letter + str(counter+2)
			sheet[name_column_value] = person_name
			book.save(output_path)


			if each_name_block.find(re.compile("title")) != None:
				position = (each_name_block.find(re.compile("title")).text)
			elif each_name_block.find(re.compile("title")) == None:
				position = "no title found for this attribute tag"
			position_column_value = position_column_letter + str(counter + 2)
			sheet[position_column_value] = position
			book.save(output_path)



			if each_name_block.find(re.compile("reportablecompfromorganization")) != None:
				reportable_comp_from_organization = (each_name_block.find(re.compile("reportablecompfromorganization")).text)
			if each_name_block.find(re.compile("reportablecompfromorgamt")) != None:
				reportable_comp_from_organization = (each_name_block.find(re.compile("reportablecompfromorgamt")).text)
			elif each_name_block.find(re.compile("reportablecompfromorganization")) == None and each_name_block.find(re.compile("reportablecompfromorgamt")) == None:
				if is_reported_compensated:
					reportable_comp_from_organization = "tag not found for this person, but present in file"
				else:
					reportable_comp_from_organization = "tag not in file"
			salary_column_value = salary_column_letter + str(counter + 2)
			sheet[salary_column_value] = reportable_comp_from_organization
			book.save(output_path)


			if each_name_block.find(re.compile("reportablecompfromrelatedorgs")) != None:
				reported_comp_from_related = (each_name_block.find(re.compile("reportablecompfromrelatedorgs")).text)
			if each_name_block.find(re.compile("reportablecompfromrltdorgamt")) != None:
				reported_comp_from_related = (each_name_block.find(re.compile("reportablecompfromrltdorgamt")).text)
			elif each_name_block.find(re.compile("reportablecompfromrelatedorgs")) == None and each_name_block.find(re.compile("reportablecompfromrltdorgamt")) == None:
				if is_related_comp:
					reported_comp_from_related = "tag not found for this person, but present in file"
				else:
					reported_comp_from_related = "tag not in file"
			extra_comp_column_value = extra_comp_column_letter + str(counter + 2)
			sheet[extra_comp_column_value] = reported_comp_from_related
			book.save(output_path)

			

			if each_name_block.find(re.compile("othercompensation")) != None:
				other_comp = (each_name_block.find(re.compile("othercompensation")).text)
			elif each_name_block.find(re.compile("othercompensation")) == None:
				if is_other:
					other_comp = "tag not found for this person, but present in file"
				else:
					other_comp = "tag not in file"
			other_comp_column_value = other_comp_column_letter + str(counter + 2)
			sheet[other_comp_column_value] = other_comp
			book.save(output_path)



			if each_name_block.find(re.compile("averagehoursperweek")) != None:
				avg_hrs_per_week = (each_name_block.find(re.compile("averagehoursperweek")).text)
			elif each_name_block.find(re.compile("averagehoursperweek")) == None:
				if is_avg_hours:
					avg_hrs_per_week = "tag not found for this person, but present in file"
				else:
					avg_hrs_per_week = "tag not in file"
			hours_per_week_column_value = hours_per_week_column_letter + str(counter + 2)
			sheet[hours_per_week_column_value] = avg_hrs_per_week
			book.save(output_path)


			if each_name_block.find(re.compile("individualtrusteeordirectorind")) != None:
				trustee_or_director = "Y"
			if each_name_block.find(re.compile("individualtrusteeordirector")) != None:
				trustee_or_director = "Y"
			elif each_name_block.find(re.compile("individualtrusteeordirector")) == None and each_name_block.find(re.compile("individualtrusteeordirectorind")) == None:
				if is_trustee_or_director:
					trustee_or_director = "N"
				else:
					trustee_or_director = "no title found for this attribute tag"
			trustee_or_director_column_value = trustee_or_director_column_letter + str(counter + 2)
			sheet[trustee_or_director_column_value] = trustee_or_director
			book.save(output_path)


			if each_name_block.find(re.compile("officer")) != None:
				officer = "Y"
			elif each_name_block.find(re.compile("officer")) == None:
				if is_officer:
					officer = "N"
				else:
					officer = "no title found for this attribute tag"
			officer_column_value = officer_column_letter + str(counter + 2)
			sheet[officer_column_value] = officer
			book.save(output_path)


			if each_name_block.find(re.compile("keyemployeeind")) != None:
				key_employee = "Y"
			if each_name_block.find(re.compile("keyemployee")) != None:
				key_employee = "Y"
			elif each_name_block.find(re.compile("keyemployeeind")) == None:
				if is_key:
					key_employee = "N"
				else:
					key_employee = "no title found for this attribute tag"
			key_employee_column_value = key_employee_column_letter + str(counter + 2)
			sheet[key_employee_column_value] = key_employee
			book.save(output_path)


			if each_name_block.find(re.compile("highestcompensatedemployee")) != None:
				highest_compensated = "Y"
			elif each_name_block.find(re.compile("highestcompensatedemployee")) == None:
				if is_highest_paid:
					highest_compensated = "N"
				else:
					highest_compensated = "no title found for this attribute tag"
			highest_paid_column_value = highest_paid_column_letter + str(counter + 2)
			sheet[highest_paid_column_value] = highest_compensated
			book.save(output_path)


			if each_name_block.find(re.compile("formerofcremployeeslistedind")) != None:
				formerly_employed = "Y"
			elif each_name_block.find(re.compile("formerofcremployeeslistedind")) == None:
				if is_former:
					formerly_employed = "N"
				else:
					formerly_employed = "no title found for this attribute tag"
			former_column_value = former_column_letter + str(counter + 2)
			sheet[former_column_value] = formerly_employed
			book.save(output_path)


			
			count_people +=1

			counter+=1
			


			#name_paragraph = each_name_block.text
			#list_of_different_lines = name_paragraph.splitlines()
			#print(list_of_different_lines)
	else:

		if xml_url==[]:
			print("NO XMLS for this ein at all!")

			ein_column_value = ein_column_letter + str(counter+2)
			sheet[ein_column_value] = ein
			book.save(output_path)

			revenue_column_value = total_revenue_column_letter + str(counter+2)
			sheet[revenue_column_value] = total_revenue
			book.save(output_path)

			name_column_value = name_column_letter + str(counter+2)
			sheet[name_column_value] = "NO XMLs for this ein at all!"
			book.save(output_path)
			
			counter+=1


		else:
			if two_filings_found_for_year==False:
				print("no xml found for this year for this ein! Try again!")

				ein_column_value = ein_column_letter + str(counter+2)
				sheet[ein_column_value] = ein
				book.save(output_path)

				revenue_column_value = total_revenue_column_letter + str(counter+2)
				sheet[revenue_column_value] = total_revenue
				book.save(output_path)

				name_column_value = name_column_letter + str(counter+2)
				sheet[name_column_value] = "could not load xml for this ein"
				book.save(output_path)

				
				counter+=1


	i+=1






