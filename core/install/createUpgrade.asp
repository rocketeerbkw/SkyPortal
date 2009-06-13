<% 
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2008 Dogg Software All Rights Reserved
'<>
'<> By using this program, you are agreeing to the terms of the
'<> SkyPortal End-User License Agreement.
'<>
'<> All copyright notices regarding SkyPortal must remain 
'<> intact in the scripts and in the outputted HTML.
'<> The "powered by" text/logo with a link back to 
'<> http://www.SkyPortal.net in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> Support can be obtained from support forums at:
'<> http://www.SkyPortal.net
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub update13x()
forumWID = 0
forumAID = 0
adminID = 0
arrWebMaster = split(strWebMaster,",")
if getmemberid(arrWebMaster(0)) <> 0 then
adminID = getmemberid(arrWebMaster(0))
else
adminID = getmemberid(arrWebMaster(1))
end if

':::::::::::::::::::::::::::::::::: CREATE PORTAL_BANNERS TABLE :::::::::::::::::::::::::::::::::::::::::
tblBanners()

':::::::::::::::::::::::::::::::::: CREATE PORTAL_EVENTS_REMINDERS TABLE :::::::::::::::::::::::::::::::::::::::::
'droptable("PORTAL_EVENTS_REMINDERS")
sSQL = "CREATE TABLE [PORTAL_EVENTS_REMINDERS]([EVENT_ID] LONG DEFAULT 0, [EVENT_START] TEXT(50), [MEMBER_ID] LONG DEFAULT 0, [REMINDER_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [REMINDER_INC] LONG DEFAULT 0);"

'createTable(checkIt(sSQL))

redim indexes(2)
indexes(0) = "CREATE INDEX [EVENT_ID] ON [PORTAL_EVENTS_REMINDERS]([EVENT_ID]);"
indexes(1) = "CREATE INDEX [MEMBER_ID] ON [PORTAL_EVENTS_REMINDERS]([MEMBER_ID]);"
indexes(2) = "CREATE INDEX [REMINDER_ID] ON [PORTAL_EVENTS_REMINDERS]([REMINDER_ID]);"
'createIndx(indexes)

':::::::::::::::::::::::::::::::::: CREATE PORTAL_EVENTS_SERIES TABLE :::::::::::::::::::::::::::::::::::::::::
'droptable("PORTAL_EVENTS_SERIES")
sSQL = "CREATE TABLE [PORTAL_EVENTS_SERIES]([SERIES_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [SERIES_TYPE] LONG DEFAULT 0, [SERIES_WEEK] LONG DEFAULT 0);"

'createTable(checkIt(sSQL))

'createIndex("CREATE INDEX [SERIES_ID] ON [PORTAL_EVENTS_SERIES]([SERIES_ID]);")
':::::::::::::::::::::::::::::::::: CREATE PORTAL_IPLIST  TABLE :::::::::::::::::::::::::::::::::::::::::
droptable("PORTAL_IPLIST")
sSQL = "CREATE TABLE [PORTAL_IPLIST]([IPLIST_COMMENT] TEXT(255), [IPLIST_DBPAGEKEY] TEXT(32), [IPLIST_ENDDATE] TEXT(32), [IPLIST_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [IPLIST_MEMBERID] TEXT(32) DEFAULT 0, [IPLIST_STARTDATE] TEXT(32), [IPLIST_STARTIP] TEXT(32), [IPLIST_STATUS] TEXT(8));"

createTable(checkIt(sSQL))

':::::::::::::::::::::::::::::::::: CREATE PORTAL_IPLOG  TABLE :::::::::::::::::::::::::::::::::::::::::
droptable("PORTAL_IPLOG")
sSQL = "CREATE TABLE [PORTAL_IPLOG]([IPLOG_DATE] TEXT(32), [IPLOG_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [IPLOG_IP] TEXT(32), [IPLOG_MEMBERID] TEXT(32) DEFAULT 0, [IPLOG_PATHINFO] TEXT(255));"

createTable(checkIt(sSQL))

':::::::::::::::::::::::::::::::::: CREATE PORTAL_PAGEKEYS TABLE :::::::::::::::::::::::::::::::::::::::::
droptable("PORTAL_PAGEKEYS")
sSQL = "CREATE TABLE [PORTAL_PAGEKEYS]([PAGEKEYS_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [PAGEKEYS_PAGEKEY] TEXT(32));"

createTable(checkIt(sSQL))

redim arrData(4)
arrData(0) = "PORTAL_PAGEKEYS"
arrData(1) = "PAGEKEYS_PAGEKEY"
arrData(2) = "'fhome.asp'"
arrData(3) = "'admin_login.asp'"
arrData(4) = "'default.asp'"
populateB(arrData)

'::::::::::::::::::: CREATE PORTAL_REPORTED_POST TABLE :::::::::::::::::::::::::::
droptable("PORTAL_REPORTED_POST")
sSQL = "CREATE TABLE [PORTAL_REPORTED_POST]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [R_ACTION_BY] LONG NOT NULL DEFAULT 0, [R_ACTION_DATE] TEXT(50) DEFAULT '0', [R_ACTION_TAKEN] MEMO DEFAULT '0', [R_COMMENTS] MEMO DEFAULT '0', [R_POST] MEMO DEFAULT '0', [R_REASON] MEMO, [R_REPLY_ID] TEXT(21) NOT NULL DEFAULT '0', [R_REPORTED_DATE] TEXT(50) DEFAULT '0', [R_REPORTER_ID] TEXT(21) NOT NULL DEFAULT '0', [R_REPORTER_IP] TEXT(20) DEFAULT '0', [R_STATUS] LONG NOT NULL DEFAULT 0, [R_TOPIC_ID] TEXT(100) NOT NULL DEFAULT '0');"

createTable(checkIt(sSQL))

':::::::::::::::::: CREATE PORTAL_COUNTRIES TABLE ::::::::::::::::::::::
redim arrCntryData(260)
arrCntryData(0) = "" & strTablePrefix & "COUNTRIES"
arrCntryData(1) = "[CO_NAME], [CO_ABBREV], [CO_CCTLD], [CO_FLAG]"
arrCntryData(2) ="'Afghanistan', 'AF', '.af', 'images/flags/af-flag.gif'"
arrCntryData(3)="'Albania', 'AL', '.al', 'images/flags/al-flag.gif'"
arrCntryData(4)="'Algeria', 'AG', '.dz', 'images/flags/ag-flag.gif'"
arrCntryData(5)="'American Samoa', 'AQ', '.as', 'images/flags/aq-flag.gif'"
arrCntryData(6)="'Andorra', 'AN', '.ad', 'images/flags/an-flag.gif'"
arrCntryData(7)="'Angola', 'AO', '.ao', 'images/flags/ao-flag.gif'"
arrCntryData(8)="'Anguilla', 'AV', '.ai', 'images/flags/av-flag.gif'"
arrCntryData(9)="'Antigua and Barbuda', 'AC', '.ag', 'images/flags/ac-flag.gif'"
arrCntryData(10)="'Argentina', 'AR', '.ar', 'images/flags/ar-flag.gif'"
arrCntryData(11)="'Armenia', 'AM', '.am', 'images/flags/am-flag.gif'"
arrCntryData(12)="'Aruba', 'AA', '.aw', 'images/flags/aa-flag.gif'"
arrCntryData(13)="'Ashmore and Cartier Islands', 'AT', '-', 'images/flags/at-flag.gif'"
arrCntryData(14)="'Australia', 'AS', '.au', 'images/flags/as-flag.gif'"
arrCntryData(15)="'Austria', 'AU', '.at', 'images/flags/au-flag.gif'"
arrCntryData(16)="'Azerbaijan', 'AJ', '.az', 'images/flags/aj-flag.gif'"
arrCntryData(17)="'Bahamas', 'BF', '.bs', 'images/flags/bf-flag.gif'"
arrCntryData(18)="'Bahrain', 'BA', '.bh', 'images/flags/ba-flag.gif'"
arrCntryData(19)="'Baker Island', 'FQ', '-', 'images/flags/fq-flag.gif'"
arrCntryData(20)="'Bangladesh', 'BG', '.bd', 'images/flags/bg-flag.gif'"
arrCntryData(21)="'Barbados', 'BB', '.bb', 'images/flags/bb-flag.gif'"
arrCntryData(22)="'Bassas da India', 'BS', '-', 'images/flags/bs-flag.gif'"
arrCntryData(23)="'Belarus', 'BO', '.by', 'images/flags/bo-flag.gif'"
arrCntryData(24)="'Belgium', 'BE', '.be', 'images/flags/be-flag.gif'"
arrCntryData(25)="'Belize', 'BH', '.bz', 'images/flags/bh-flag.gif'"
arrCntryData(26)="'Benin', 'BN', '.bj', 'images/flags/bn-flag.gif'"
arrCntryData(27)="'Bermuda', 'BD', '.bm', 'images/flags/bd-flag.gif'"
arrCntryData(28)="'Bhutan', 'BT', '.bt', 'images/flags/bt-flag.gif'"
arrCntryData(29)="'Bolivia', 'BL', '.bo', 'images/flags/bl-flag.gif'"
arrCntryData(30)="'Bosnia and Herzegovina', 'BK', '.ba', 'images/flags/bk-flag.gif'"
arrCntryData(31)="'Botswana', 'BC', '.bw', 'images/flags/bc-flag.gif'"
arrCntryData(32)="'Bouvet Island', 'BV', '.bv', 'images/flags/bv-flag.gif'"
arrCntryData(33)="'Brazil', 'BR', '.br', 'images/flags/br-flag.gif'"
arrCntryData(34)="'British Indian Ocean Territories', 'IO', '.io', 'images/flags/io-flag.gif'"
arrCntryData(35)="'Brunei', 'BX', '.bn', 'images/flags/bx-flag.gif'"
arrCntryData(36)="'Bulgaria', 'BU', '.bg', 'images/flags/bu-flag.gif'"
arrCntryData(37)="'Burkina Faso (Upper Volta)', 'UV', '.bf', 'images/flags/uv-flag.gif'"
arrCntryData(38)="'Burundi', 'BY', '.bi', 'images/flags/by-flag.gif'"
arrCntryData(39)="'Cambodia', 'CB', '.kh', 'images/flags/cb-flag.gif'"
arrCntryData(40)="'Cameroon', 'CM', '.cm', 'images/flags/cm-flag.gif'"
arrCntryData(41)="'Canada', 'CA', '.ca', 'images/flags/ca-flag.gif'"
arrCntryData(42)="'Cape Vere Islands', 'CV', '.cv', 'images/flags/cv-flag.gif'"
arrCntryData(43)="'Cayman Islands', 'CJ', '.ky', 'images/flags/cj-flag.gif'"
arrCntryData(44)="'Central African Republic', 'CT', '.cf', 'images/flags/ct-flag.gif'"
arrCntryData(45)="'Chad', 'CD', '.td', 'images/flags/cd-flag.gif'"
arrCntryData(46)="'Chile', 'CI', '.cl', 'images/flags/ci-flag.gif'"
arrCntryData(47)="'China', 'CH', '.cn', 'images/flags/ch-flag.gif'"
arrCntryData(48)="'Christmas Island', 'KT', '.cx', 'images/flags/kt-flag.gif'"
arrCntryData(49)="'Clipperton Island', 'IP', '-', 'images/flags/ip-flag.gif'"
arrCntryData(50)="'Cocos (Keeling) Islands', 'CK', '.cc', 'images/flags/ck-flag.gif'"
arrCntryData(51)="'Colombia', 'CO', '.co', 'images/flags/co-flag.gif'"
arrCntryData(52)="'Comoros Islands', 'CN', '.km', 'images/flags/cn-flag.gif'"
arrCntryData(53)="'Congo, Democratic Republic of', 'CG', '.cd', 'images/flags/cg-flag.gif'"
arrCntryData(54)="'Congo, Republic of the', 'CF', '.cg', 'images/flags/cf-flag.gif'"
arrCntryData(55)="'Cook Islands', 'CW', '.ck', 'images/flags/cw-flag.gif'"
arrCntryData(56)="'Coral Sea Islands', 'CR', '-', 'images/flags/cr-flag.gif'"
arrCntryData(57)="'Costa Rica', 'CS', '.cr', 'images/flags/cs-flag.gif'"
arrCntryData(58)="'Cote d Ivoire', 'IV', '.ci', 'images/flags/iv-flag.gif'"
arrCntryData(59)="'Croatia', 'HR', '.hr', 'images/flags/hr-flag.gif'"
arrCntryData(60)="'Cuba', 'CU', '.cu', 'images/flags/cu-flag.gif'"
arrCntryData(61)="'Cyprus', 'CY', '.cy', 'images/flags/cy-flag.gif'"
arrCntryData(62)="'Czech Republic', 'EZ', '.cz', 'images/flags/ez-flag.gif'"
arrCntryData(63)="'Denmark', 'DA', '.dk', 'images/flags/da-flag.gif'"
arrCntryData(64)="'Djibouti', 'DJ', '.dj', 'images/flags/dj-flag.gif'"
arrCntryData(65)="'Dominica', 'DO', '.dm', 'images/flags/do-flag.gif'"
arrCntryData(66)="'Dominican Republic', 'DR', '.do', 'images/flags/dr-flag.gif'"
arrCntryData(67)="'East Timor', 'TT', '.tp', 'images/flags/tt-flag.gif'"
arrCntryData(68)="'Ecuador', 'EC', '.ec', 'images/flags/ec-flag.gif'"
arrCntryData(69)="'Egypt', 'EG', '.eg', 'images/flags/eg-flag.gif'"
arrCntryData(70)="'El Salvador', 'ES', '.sv', 'images/flags/es-flag.gif'"
arrCntryData(71)="'Equatorial Guinea', 'EK', '.gq', 'images/flags/ek-flag.gif'"
arrCntryData(72)="'Eritrea', 'ER', '.er', 'images/flags/er-flag.gif'"
arrCntryData(73)="'Estonia', 'EN', '.ee', 'images/flags/en-flag.gif'"
arrCntryData(74)="'Ethiopia', 'ET', '.et', 'images/flags/et-flag.gif'"
arrCntryData(75)="'Europa Island', 'EU', '-', 'images/flags/eu-flag.gif'"
arrCntryData(76)="'Falkland Islands', 'FK', '.fk', 'images/flags/fk-flag.gif'"
arrCntryData(77)="'Faroe Islands', 'FO', '.fo', 'images/flags/fo-flag.gif'"
arrCntryData(78)="'Fiji', 'FJ', '.fj', 'images/flags/fj-flag.gif'"
arrCntryData(79)="'Finland', 'FI', '.fi', 'images/flags/fi-flag.gif'"
arrCntryData(80)="'France', 'FR', '.fr', 'images/flags/fr-flag.gif'"
arrCntryData(81)="'French Guiana', 'FG', '.gf', 'images/flags/fg-flag.gif'"
arrCntryData(82)="'French Polynesia', 'FP', '.pf', 'images/flags/fp-flag.gif'"
arrCntryData(83)="'French Southern and Antarctic Lands', 'FS', '.tf', 'images/flags/fs-flag.gif'"
arrCntryData(84)="'Fyro Macedonia', 'MK', '.mk', 'images/flags/mk-flag.gif'"
arrCntryData(85)="'Gabon', 'GB', '.ga', 'images/flags/gb-flag.gif'"
arrCntryData(86)="'Gambia', 'GA', '.gm', 'images/flags/ga-flag.gif'"
arrCntryData(87)="'Georgia', 'GG', '.ge', 'images/flags/gg-flag.gif'"
arrCntryData(88)="'Germany', 'GM', '.de', 'images/flags/gm-flag.gif'"
arrCntryData(89)="'Ghana', 'GH', '.gh', 'images/flags/gh-flag.gif'"
arrCntryData(90)="'Gibraltar', 'GI', '.gi', 'images/flags/gi-flag.gif'"
arrCntryData(91)="'Glorioso Islands', 'GO', '-', 'images/flags/go-flag.gif'"
arrCntryData(92)="'Greece', 'GR', '.gr', 'images/flags/gr-flag.gif'"
arrCntryData(93)="'Greenland', 'GL', '.gl', 'images/flags/gl-flag.gif'"
arrCntryData(94)="'Grenada', 'GJ', '.gd', 'images/flags/gj-flag.gif'"
arrCntryData(95)="'Guadeloupe', 'GP', '.gp', 'images/flags/gp-flag.gif'"
arrCntryData(96)="'Guam', 'GQ', '.gu', 'images/flags/gq-flag.gif'"
arrCntryData(97)="'Guatemala', 'GT', '.gt', 'images/flags/gt-flag.gif'"
arrCntryData(98)="'Guernsey', 'GK', '.gg', 'images/flags/gk-flag.gif'"
arrCntryData(99)="'Guinea', 'GV', '.gn', 'images/flags/gv-flag.gif'"
arrCntryData(100)="'Guinea-Bissau', 'PU', '.gw', 'images/flags/pu-flag.gif'"
arrCntryData(101)="'Guyana', 'GY', '.gy', 'images/flags/gy-flag.gif'"
arrCntryData(102)="'Haiti', 'HA', '.ht', 'images/flags/ha-flag.gif'"
arrCntryData(103)="'Heard Island and McDonald Islands', 'HM', '.hm', 'images/flags/hm-flag.gif'"
arrCntryData(104)="'Holy See (Vatican City)', 'VT', '.va', 'images/flags/vt-flag.gif'"
arrCntryData(105)="'Honduras', 'HO', '.hn', 'images/flags/ho-flag.gif'"
arrCntryData(106)="'Hong Kong', 'HK', '.hk', 'images/flags/hk-flag.gif'"
arrCntryData(107)="'Howland Island', 'HQ', '-', 'images/flags/hq-flag.gif'"
arrCntryData(108)="'Hungary', 'HU', '.hu', 'images/flags/hu-flag.gif'"
arrCntryData(109)="'Iceland', 'IC', '.is', 'images/flags/ic-flag.gif'"
arrCntryData(110)="'India', 'IN', '.in', 'images/flags/in-flag.gif'"
arrCntryData(111)="'Indonesia', 'ID', '.id', 'images/flags/id-flag.gif'"
arrCntryData(112)="'Iran', 'IR', '.ir', 'images/flags/ir-flag.gif'"
arrCntryData(113)="'Iraq', 'IZ', '.iq', 'images/flags/iz-flag.gif'"
arrCntryData(114)="'Ireland', 'EI', '.ie', 'images/flags/ei-flag.gif'"
arrCntryData(115)="'Israel', 'IS', '.il', 'images/flags/is-flag.gif'"
arrCntryData(116)="'Italy', 'IT', '.it', 'images/flags/it-flag.gif'"
arrCntryData(117)="'Jamaica', 'JM', '.jm', 'images/flags/jm-flag.gif'"
arrCntryData(118)="'Jan Mayen', 'JN', '-', 'images/flags/jn-flag.gif'"
arrCntryData(119)="'Japan', 'JA', '.jp', 'images/flags/ja-flag.gif'"
arrCntryData(120)="'Jarvis Island', 'DQ', '-', 'images/flags/dq-flag.gif'"
arrCntryData(121)="'Jersey', 'JE', '.je', 'images/flags/je-flag.gif'"
arrCntryData(122)="'Johnston Atoll', 'JQ', '-', 'images/flags/jq-flag.gif'"
arrCntryData(123)="'Jordan', 'JO', '.jo', 'images/flags/jo-flag.gif'"
arrCntryData(124)="'Juan de Nova Island', 'JU', '-', 'images/flags/ju-flag.gif'"
arrCntryData(125)="'Kazakhstan', 'KZ', '.kz', 'images/flags/kz-flag.gif'"
arrCntryData(126)="'Kenya', 'KE', '.ke', 'images/flags/ke-flag.gif'"
arrCntryData(127)="'Kingman Reef', 'KQ', '-', 'images/flags/kq-flag.gif'"
arrCntryData(128)="'Kiribati', 'KR', '.ki', 'images/flags/kr-flag.gif'"
arrCntryData(129)="'Korea, North', 'KN', '.kp', 'images/flags/kn-flag.gif'"
arrCntryData(130)="'Korea, South', 'KS', '.kr', 'images/flags/ks-flag.gif'"
arrCntryData(131)="'Kuwait', 'KU', '.kw', 'images/flags/ku-flag.gif'"
arrCntryData(132)="'Kyrgyzstan', 'KG', '.kg', 'images/flags/kg-flag.gif'"
arrCntryData(133)="'Laos', 'LA', '.la', 'images/flags/la-flag.gif'"
arrCntryData(134)="'Latvia', 'LG', '.lv', 'images/flags/lg-flag.gif'"
arrCntryData(135)="'Lebanon', 'LE', '.lb', 'images/flags/le-flag.gif'"
arrCntryData(136)="'Lesotho', 'LT', '.ls', 'images/flags/lt-flag.gif'"
arrCntryData(137)="'Liberia', 'LI', '.lr', 'images/flags/li-flag.gif'"
arrCntryData(138)="'Libya', 'LY', '.ly', 'images/flags/ly-flag.gif'"
arrCntryData(139)="'Liechtenstein', 'LS', '.li', 'images/flags/ls-flag.gif'"
arrCntryData(140)="'Lithuania', 'LH', '.lt', 'images/flags/lh-flag.gif'"
arrCntryData(141)="'Luxembourg', 'LU', '.lu', 'images/flags/lu-flag.gif'"
arrCntryData(142)="'Macau', 'MC', '.mo', 'images/flags/mc-flag.gif'"
arrCntryData(143)="'Madagascar', 'MA', '.mg', 'images/flags/ma-flag.gif'"
arrCntryData(144)="'Malawi', 'MI', '.mw', 'images/flags/mi-flag.gif'"
arrCntryData(145)="'Malaysia', 'MY', '.my', 'images/flags/my-flag.gif'"
arrCntryData(146)="'Maldives', 'MV', '.mv', 'images/flags/mv-flag.gif'"
arrCntryData(147)="'Mali', 'ML', '.ml', 'images/flags/ml-flag.gif'"
arrCntryData(148)="'Malta', 'MT', '.mt', 'images/flags/mt-flag.gif'"
arrCntryData(149)="'Man, Isle of', 'IM', '.im', 'images/flags/im-flag.gif'"
arrCntryData(150)="'Marshall Islands', 'RM', '.mh', 'images/flags/rm-flag.gif'"
arrCntryData(151)="'Martinique', 'MB', '.mq', 'images/flags/mb-flag.gif'"
arrCntryData(152)="'Mauritania', 'MR', '.mr', 'images/flags/mr-flag.gif'"
arrCntryData(153)="'Mauritius', 'MP', '.mu', 'images/flags/mp-flag.gif'"
arrCntryData(154)="'Mayotte', 'MF', '.yt', 'images/flags/mf-flag.gif'"
arrCntryData(155)="'Mexico', 'MX', '.mx', 'images/flags/mx-flag.gif'"
arrCntryData(156)="'Micronesia, Federated States of', 'FM', '.fm', 'images/flags/fm-flag.gif'"
arrCntryData(157)="'Midway Islands', 'MQ', '-', 'images/flags/mq-flag.gif'"
arrCntryData(158)="'Moldova', 'MD', '.md', 'images/flags/md-flag.gif'"
arrCntryData(159)="'Monaco', 'MN', '.mc', 'images/flags/mn-flag.gif'"
arrCntryData(160)="'Mongolia', 'MG', '.mn', 'images/flags/mg-flag.gif'"
arrCntryData(161)="'Montserrat', 'MH', '.ms', 'images/flags/mh-flag.gif'"
arrCntryData(162)="'Morocco', 'MO', '.ma', 'images/flags/mo-flag.gif'"
arrCntryData(163)="'Mozambique', 'MZ', '.mz', 'images/flags/mz-flag.gif'"
arrCntryData(164)="'Myanmar (Burma)', 'BM', '.mm', 'images/flags/bm-flag.gif'"
arrCntryData(165)="'Namibia', 'WA', '.na', 'images/flags/wa-flag.gif'"
arrCntryData(166)="'Nauru', 'NR', '.nr', 'images/flags/nr-flag.gif'"
arrCntryData(167)="'Navassa Island', 'BQ', '-', 'images/flags/bq-flag.gif'"
arrCntryData(168)="'Nepal', 'NP', '.np', 'images/flags/np-flag.gif'"
arrCntryData(169)="'Netherlands', 'NL', '.nl', 'images/flags/nl-flag.gif'"
arrCntryData(170)="'Netherlands Antilles', 'NT', '.an', 'images/flags/nt-flag.gif'"
arrCntryData(171)="'New Caledonia', 'NC', '.nc', 'images/flags/nc-flag.gif'"
arrCntryData(172)="'New Zealand', 'NZ', '.nz', 'images/flags/nz-flag.gif'"
arrCntryData(173)="'Nicaragua', 'NU', '.ni', 'images/flags/nu-flag.gif'"
arrCntryData(174)="'Niger', 'NG', '.ne', 'images/flags/ng-flag.gif'"
arrCntryData(175)="'Nigeria', 'NI', '.ng', 'images/flags/ni-flag.gif'"
arrCntryData(176)="'Niue', 'NE', '.nu', 'images/flags/ne-flag.gif'"
arrCntryData(177)="'Norfolk Island', 'NF', '.nf', 'images/flags/nf-flag.gif'"
arrCntryData(178)="'Northern Mariana Islands', 'CQ', '.mp', 'images/flags/cq-flag.gif'"
arrCntryData(179)="'Norway', 'NO', '.no', 'images/flags/no-flag.gif'"
arrCntryData(180)="'Oman', 'MU', '.om', 'images/flags/mu-flag.gif'"
arrCntryData(181)="'Pakistan', 'PK', '.pk', 'images/flags/pk-flag.gif'"
arrCntryData(182)="'Palau', 'PS', '.pw', 'images/flags/ps-flag.gif'"
arrCntryData(183)="'Palmyra Atoll', 'LQ', '-', 'images/flags/lq-flag.gif'"
arrCntryData(184)="'Panama', 'PM', '.pa', 'images/flags/pm-flag.gif'"
arrCntryData(185)="'Papua New Guinea', 'PP', '.pg', 'images/flags/pp-flag.gif'"
arrCntryData(186)="'Paracel Islands', 'PF', '-', 'images/flags/cw-flag.gif'"
arrCntryData(187)="'Paraguay', 'PA', '.py', 'images/flags/pa-flag.gif'"
arrCntryData(188)="'Peru', 'PE', '.pe', 'images/flags/pe-flag.gif'"
arrCntryData(189)="'Philippines', 'RP', '.ph', 'images/flags/rp-flag.gif'"
arrCntryData(190)="'Pitcairn Island', 'PC', '.pn', 'images/flags/pc-flag.gif'"
arrCntryData(191)="'Poland', 'PL', '.pl', 'images/flags/pl-flag.gif'"
arrCntryData(192)="'Portugal', 'PO', '.pt', 'images/flags/po-flag.gif'"
arrCntryData(193)="'Puerto Rico', 'RQ', '.pr', 'images/flags/rq-flag.gif'"
arrCntryData(194)="'Qatar', 'QA', '.qa', 'images/flags/qa-flag.gif'"
arrCntryData(195)="'Reunion Island', 'RE', '.re', 'images/flags/re-flag.gif'"
arrCntryData(196)="'Romania', 'RO', '.ro', 'images/flags/ro-flag.gif'"
arrCntryData(197)="'Russia', 'RS', '.ru', 'images/flags/rs-flag.gif'"
arrCntryData(198)="'Rwanda', 'RW', '.rw', 'images/flags/rw-flag.gif'"
arrCntryData(199)="'Saint Helena', 'SH', '.sh', 'images/flags/sh-flag.gif'"
arrCntryData(200)="'Saint Kitts and Nevis', 'SC', '.kn', 'images/flags/sc-flag.gif'"
arrCntryData(201)="'Saint Lucia', 'ST', '.lc', 'images/flags/st-flag.gif'"
arrCntryData(202)="'Saint Pierre and Miquelon', 'SB', '.pm', 'images/flags/sb-flag.gif'"
arrCntryData(203)="'Saint Vincent and the Grenadines', 'VC', '.vc', 'images/flags/vc-flag.gif'"
arrCntryData(204)="'Samoa', 'WS', '.ws', 'images/flags/ws-flag.gif'"
arrCntryData(205)="'San Marino', 'SM', '.sm', 'images/flags/sm-flag.gif'"
arrCntryData(206)="'Sao Tome and Principe', 'TP', '.st', 'images/flags/tp-flag.gif'"
arrCntryData(207)="'Saudi Arabia', 'SA', '.sa', 'images/flags/sa-flag.gif'"
arrCntryData(208)="'Senegal', 'SG', '.sn', 'images/flags/sg-flag.gif'"
arrCntryData(209)="'Serbia and Montenegro', 'YI', '.yu', 'images/flags/yi-flag.gif'"
arrCntryData(210)="'Seychelles', 'SE', '.sc', 'images/flags/se-flag.gif'"
arrCntryData(211)="'Sierra Leone', 'SL', '.sl', 'images/flags/sl-flag.gif'"
arrCntryData(212)="'Singapore', 'SN', '.sg', 'images/flags/sn-flag.gif'"
arrCntryData(213)="'Slovakia', 'LO', '.sk', 'images/flags/lo-flag.gif'"
arrCntryData(214)="'Slovenia', 'SI', '.si', 'images/flags/si-flag.gif'"
arrCntryData(215)="'Solomon Islands', 'BP', '.sb', 'images/flags/bp-flag.gif'"
arrCntryData(216)="'Somalia', 'SO', '.so', 'images/flags/so-flag.gif'"
arrCntryData(217)="'South Africa', 'SF', '.za', 'images/flags/sf-flag.gif'"
arrCntryData(218)="'South Georgia and the Islands', 'SX', '.gs', 'images/flags/sx-flag.gif'"
arrCntryData(219)="'Spain', 'SP', '.es', 'images/flags/sp-flag.gif'"
arrCntryData(220)="'Spratly Islands', 'PG', '-', 'images/flags/no-flag.gif'"
arrCntryData(221)="'Sri Lanka', 'CE', '.lk', 'images/flags/ce-flag.gif'"
arrCntryData(222)="'Sudan', 'SU', '.sd', 'images/flags/su-flag.gif'"
arrCntryData(223)="'Suriname', 'NS', '.sr', 'images/flags/ns-flag.gif'"
arrCntryData(224)="'Svalbard', 'SV', '.sj', 'images/flags/sv-flag.gif'"
arrCntryData(225)="'Swaziland', 'WZ', '.sz', 'images/flags/wz-flag.gif'"
arrCntryData(226)="'Sweden', 'SW', '.se', 'images/flags/sw-flag.gif'"
arrCntryData(227)="'Switzerland', 'SZ', '.ch', 'images/flags/sz-flag.gif'"
arrCntryData(228)="'Syria', 'SY', '.sy', 'images/flags/sy-flag.gif'"
arrCntryData(229)="'Taiwan', 'TW', '.tw', 'images/flags/tw-flag.gif'"
arrCntryData(230)="'Tajikistan', 'TI', '.tj', 'images/flags/ti-flag.gif'"
arrCntryData(231)="'Tanzania', 'TZ', '.tz', 'images/flags/tz-flag.gif'"
arrCntryData(232)="'Thailand', 'TH', '.th', 'images/flags/th-flag.gif'"
arrCntryData(233)="'Togo', 'TO', '.tg', 'images/flags/to-flag.gif'"
arrCntryData(234)="'Tokelau', 'TL', '.tk', 'images/flags/tl-flag.gif'"
arrCntryData(235)="'Tonga', 'TN', '.to', 'images/flags/tn-flag.gif'"
arrCntryData(236)="'Trinidad and Tobago', 'TD', '.tt', 'images/flags/td-flag.gif'"
arrCntryData(237)="'Tromelin Island', 'TE', '-', 'images/flags/te-flag.gif'"
arrCntryData(238)="'Tunisia', 'TS', '.tn', 'images/flags/ts-flag.gif'"
arrCntryData(239)="'Turkey', 'TU', '.tr', 'images/flags/tu-flag.gif'"
arrCntryData(240)="'Turkmenistan', 'TX', '.tm', 'images/flags/tx-flag.gif'"
arrCntryData(241)="'Turks and Caicos Islands', 'TK', '.tc', 'images/flags/tk-flag.gif'"
arrCntryData(242)="'Tuvalu', 'TV', '.tv', 'images/flags/tv-flag.gif'"
arrCntryData(243)="'Uganda', 'UG', '.ug', 'images/flags/ug-flag.gif'"
arrCntryData(244)="'Ukraine', 'UP', '.ua', 'images/flags/up-flag.gif'"
arrCntryData(245)="'United Arab Emirates', 'AE', '.ae', 'images/flags/ae-flag.gif'"
arrCntryData(246)="'United Kingdom', 'UK', '.uk', 'images/flags/uk-flag.gif'"
arrCntryData(247)="'Uruguay', 'UY', '.uy', 'images/flags/uy-flag.gif'"
arrCntryData(248)="'USA', 'US', '.us', 'images/flags/us-flag.gif'"
arrCntryData(249)="'Uzbekistan', 'UZ', '.uz', 'images/flags/uz-flag.gif'"
arrCntryData(250)="'Vanuatu', 'NH', '.vu', 'images/flags/nh-flag.gif'"
arrCntryData(251)="'Venezuela', 'VE', '.ve', 'images/flags/ve-flag.gif'"
arrCntryData(252)="'Vietnam', 'VM', '.vn', 'images/flags/vm-flag.gif'"
arrCntryData(253)="'Virgin Islands (UK)', 'VI', '.vg', 'images/flags/vi-flag.gif'"
arrCntryData(254)="'Virgin Islands (US)', 'VQ', '.vi', 'images/flags/vq-flag.gif'"
arrCntryData(255)="'Wake Island', 'WQ', '-', 'images/flags/wq-flag.gif'"
arrCntryData(256)="'Wallis and Futuna Islands ', 'WF', '.wf', 'images/flags/wf-flag.gif'"
arrCntryData(257)="'Western Sahara', 'WI', '.eh', 'images/flags/wi-flag.gif'"
arrCntryData(258)="'Yemen', 'YM', '.ye', 'images/flags/ym-flag.gif'"
arrCntryData(259)="'Zambia', 'ZA', '.zm', 'images/flags/za-flag.gif'"
arrCntryData(260)="'Zimbabwe (Rhodesia)', 'ZI', '.zw', 'images/flags/zi-flag.gif'"

tblCountries()

'::::::::::::::::: UPDATE COUNTRIES IN MEMBERS PROFILE FOR CHANGES ::::::::::::::::::::::::::::::
' ---- Only needed for upgrade of existing db that used the old countries list (everthing up to beta3)

strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Portugal' WHERE M_COUNTRY='Azores';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Vanuatu' WHERE M_COUNTRY='Borneo';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Cameroon' WHERE M_COUNTRY='Camaroon';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Spain' WHERE M_COUNTRY='Canary Islands';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Cayman Islands' WHERE M_COUNTRY='Cayman Island';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Central African Republic' WHERE M_COUNTRY='Central African Rep';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Eritrea' WHERE M_COUNTRY='Eritria';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Fyro Macedonia' WHERE M_COUNTRY='Fed Rep Yugoslavia';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Cote d Ivoire' WHERE M_COUNTRY='Ivory Coast';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Korea, North' WHERE M_COUNTRY='Korea';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Macau' WHERE M_COUNTRY='Macao';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Nauru' WHERE M_COUNTRY='Naura';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Korea, South' WHERE M_COUNTRY='Republic of Korea';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Guadeloupe' WHERE M_COUNTRY='Saint Barthelemy';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Virgin Islands (US)' WHERE M_COUNTRY='Saint Croix';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Saint Vincent and the Grenadines' WHERE M_COUNTRY='Saint Vincent and Grenadi';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Somalia' WHERE M_COUNTRY='Somalia Northern Region';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Somalia' WHERE M_COUNTRY='Somalia Southern Region ';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'South Georgia and the Islands' WHERE M_COUNTRY='South Sandwich Islands';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Turks and Caicos Islands' WHERE M_COUNTRY='Turks and Caicos Islnd';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Virgin Islands (UK)' WHERE M_COUNTRY='Virgin Islands (United Kingdom)';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Samoa' WHERE M_COUNTRY='Western Samoa';"
doSQL2 strSql,0
strSQL="UPDATE PORTAL_MEMBERS SET M_COUNTRY = 'Congo, Democratic Republic of' WHERE M_COUNTRY='Zaire';"
doSQL2 strSql,0
Response.Write("Existing Members successfully updated for Country Changes<br />")
Response.Write("Country/Flag upgrade completed")

'::::::::::::::::::::::::::::::::::: adjust portal_config :::::::::::::::::::::::::::::::::::::::
strSQL="UPDATE PORTAL_CONFIG SET C_STREMAILVAL = 5;"
doSQL2 strSql,0

':::::::::::::::::::::::::::::::::: ALTER PORTAL_CONFIG TABLE add fields :::::::::::::::::::::::::::::::::::::::::

		response.Write("<hr /><hr />")

'	strSql = "ALTER TABLE PORTAL_CONFIG ADD"
	strSql = "PORTAL_CONFIG"
	strSql = strSql & ",C_STRDEFTHEME TEXT(50)"
	strSql = strSql & ",C_PMTYPE BYTE"
	strSql = strSql & ",C_PAGEWIDTH TEXT(50)"
	strSql = strSql & ",C_ALLOWUPLOADS BYTE"
	strSql = strSql & ",C_STRICSLOCATION TEXT(50)"
	strSql = strSql & ",C_REMINDERS LONG"
	strSql = strSql & ",C_ICALEXIST LONG"
	strSql = strSql & ",C_ICALNEW LONG"
	strSql = strSql & ",C_STRVAR1 TEXT(100)"
	strSql = strSql & ",C_STRVAR2 TEXT(100)"
	strSql = strSql & ",C_STRVAR3 TEXT(100)"
	strSql = strSql & ",C_STRVAR4 TEXT(100)"
	strSql = strSql & ",C_FORUMSUBSCRIPTION LONG"
	strSql = strSql & ",AUTOPM_ON LONG"
	strSql = strSql & ",AUTOPM_SUBJECTLINE TEXT(255)"
	strSql = strSql & ",AUTOPM_MESSAGE MEMO"
	strSql = strSql & ",C_STRZIP BYTE"
	strSql = strSql & ",C_STRMSN BYTE"
	strSql = strSql & ",C_STRIPGATEBAN TEXT(2)"
	strSql = strSql & ",C_STRIPGATECSS TEXT(2)"
	strSql = strSql & ",C_STRIPGATECOK TEXT(2)"
	strSql = strSql & ",C_STRIPGATEMET TEXT(2)"
	strSql = strSql & ",C_STRIPGATEMSG TEXT(100)"
	strSql = strSql & ",C_STRIPGATELKMSG TEXT(100)"
	strSql = strSql & ",C_STRIPGATENOACMSG TEXT(100)"
	strSql = strSql & ",C_STRIPGATEWARNMSG TEXT(100)"
	strSql = strSql & ",C_STRIPGATEVER TEXT(15)"
	strSql = strSql & ",C_STRIPGATELOG TEXT(2)"
	strSql = strSql & ",C_STRIPGATETYP TEXT(2)"
	strSql = strSql & ",C_STRIPGATEEXP TEXT(3)"
	strSql = strSql & ",C_STRIPGATELCK TEXT(2)"
	strSql = strSql & ",C_STRLOGINTYPE BYTE"
	strSql = strSql & ",C_STRLOCKDOWN BYTE"
	strSql = strSql & ",C_STRGLOW BYTE"
	strSql = strSql & ",C_STRHEADERTYPE LONG"
	strSql = strSql & ",C_MODULES TEXT(50) NULL"
		
		alterTable2(checkIt(strSql))

	'-------------------- populate table with default values --------------------------
		strSql = "UPDATE PORTAL_CONFIG SET"
		strSql = strSql & " C_STRTITLEIMAGE = 'site_Logo.jpg'"
		strSql = strSql & ", C_STRDEFAULTFONTFACE = 'Verdana, Arial, Helvetica'"
		strSql = strSql & ", C_STRDEFAULTFONTSIZE = 2"
		strSql = strSql & ", C_STRHEADERFONTSIZE = 4"
		strSql = strSql & ", C_STRFOOTERFONTSIZE = 1"
		strSql = strSql & ", C_STRDEFAULTFONTCOLOR = '#191970'"
		strSql = strSql & ", C_STRHEADCELLCOLOR = '#99B2D1'"
		strSql = strSql & ", C_STRALTHEADCELLCOLOR = '#DAE2EF'"
		strSql = strSql & ", C_STRHEADFONTCOLOR = '#ffffff'"
		strSql = strSql & ", C_STRCATEGORYCELLCOLOR = '#ABBDDC'"
		strSql = strSql & ", C_STRCATEGORYFONTCOLOR = '#FF4500'"
		strSql = strSql & ", C_STRFORUMFIRSTCELLCOLOR = '#E7E7EA'"
		strSql = strSql & ", C_STRFORUMCELLCOLOR = '#E7E7EA'"
		strSql = strSql & ", C_STRALTFORUMCELLCOLOR = '#F1F1F4'"
		strSql = strSql & ", C_STRFORUMFONTCOLOR = '#191970'"
		strSql = strSql & ", C_STRFORUMLINKCOLOR = '#00008B'"
		strSql = strSql & ", C_STRTABLEBORDERCOLOR = '#000000'"
		strSql = strSql & ", C_STRPOPUPTABLECOLOR = '#C4D1E6'"
		strSql = strSql & ", C_STRPOPUPBORDERCOLOR = '#000000'"
		strSql = strSql & ", C_STRNEWFONTCOLOR = '#0000FF'"
		strSql = strSql & ", C_STRTOPICWIDTHLEFT = '120'"
		strSql = strSql & ", C_STRTOPICNOWRAPLEFT = 1"
		strSql = strSql & ", C_STRTOPICWIDTHRIGHT = '100%'"
		strSql = strSql & ", C_STRTOPICNOWRAPRIGHT = 1"
		strSql = strSql & ", C_STRPAGEBGCOLOR = '#C6C9D1'"
		strSql = strSql & ", C_STRPAGEBGIMAGE = 'background.gif'"
		strSql = strSql & ", C_STRHEADERTYPE = 2"
		strSql = strSql & ", C_STRDEFTHEME = 'Harmonia'"
		strSql = strSql & ", C_PMTYPE = 2"
		strSql = strSql & ", C_PAGEWIDTH = '100'"
		strSql = strSql & ", C_ALLOWUPLOADS = 1"
		strSql = strSql & ", C_STRICSLOCATION = 'files/eventfile.ics'"
		strSql = strSql & ", C_REMINDERS = 1"
		strSql = strSql & ", C_ICALEXIST = 1"
		strSql = strSql & ", C_ICALNEW = 1"
		strSql = strSql & ", C_STRVAR1 = 'Shoe Size'"
		strSql = strSql & ", C_STRVAR2 = 'Favorite Foods'"
		strSql = strSql & ", C_STRVAR3 = 'My Car is a'"
		strSql = strSql & ", C_STRVAR4 = 'PacMan High Score'"
		strSql = strSql & ", C_FORUMSUBSCRIPTION = 1"
		strSql = strSql & ", AUTOPM_ON = 1"
		strSql = strSql & ", AUTOPM_SUBJECTLINE = 'Welcome!'"
		strSql = strSql & ", AUTOPM_MESSAGE = 'Welcome to our site'"
		strSql = strSql & ", C_STRZIP = 1"
		strSql = strSql & ", C_STRMSN = 1"
		strSql = strSql & ", C_STRIPGATEBAN = '0'"
		strSql = strSql & ", C_STRIPGATELCK = '0'"
		strSql = strSql & ", C_STRIPGATECOK = '1'"
		strSql = strSql & ", C_STRIPGATEMET = '1'"
		strSql = strSql & ", C_STRIPGATEMSG = 'You are banned from this site'"
		strSql = strSql & ", C_STRIPGATELKMSG = 'Forums are currently locked'"
		strSql = strSql & ", C_STRIPGATENOACMSG = 'You do not have access to the requested page'"
'		strSql = strSql & ", C_STRIPGATEWARNMSG = ''"
		strSql = strSql & ", C_STRIPGATEVER = 'Ver 2.3.0.MWP'"
		strSql = strSql & ", C_STRIPGATELOG = '1'"
		strSql = strSql & ", C_STRIPGATETYP = '0'"
		strSql = strSql & ", C_STRIPGATEEXP = '15'"
		strSql = strSql & ", C_STRIPGATECSS = '2'"
		strSql = strSql & ", C_STRLOGINTYPE = 1"
		strSql = strSql & ", C_STRLOCKDOWN = 0"
		strSql = strSql & ", C_STRGLOW = 1"
		strSql = strSql & ", C_STRGFXBUTTONS = 1"
		strSql = strSql & ", C_MODULES = ':1:2:3:4:5:6:7:8:'"
		strSql = strSql & " WHERE CONFIG_ID = 1"
'		response.Write(strSql)
on error resume next
		doSQL2 strSql,0
		

	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>Table data updated succesfully</b><br /><br />" & vbNewLine)
	end if
	err.clear
		
		response.Write("PORTAL_CONFIG")
		response.Write("<hr /><hr />")
		
':::::::::::::::::::::::::::::::::: ALTER PORTAL_CP_CONFIG TABLE :::::::::::::::::::::::::::::::::::::::::
		response.Write("<hr /><hr />PORTAL_CP_CONFIG")
	if strDBType = "sqlserver" then
	strSql = "alter table [PORTAL_CP_CONFIG] drop Constraint [PK_PORTAL_CP_CONFIG];alter table [PORTAL_CP_CONFIG] drop Constraint [DF__PORTAL_CP__THEME__5224328E];"
	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	  else
	  end if
	else
		Response.Write("    <b>Constraint dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
		
	strSql = "ALTER TABLE [PORTAL_CP_CONFIG] ADD [ID] int IDENTITY CONSTRAINT [ID] PRIMARY KEY"
	my_Conn.execute(checkIt(strSql))
	if err.number <> 0 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>Table PORTAL_CP_CONFIG primary key: ID added</b><br /><br />" & vbNewLine)
	end if
	err.clear
	
':::::::::::::::::::::::::::::::::: ALTER PORTAL_ARCHIVE_REPLY TABLE :::::::::::::::::::::::::::::::::::::::::
	response.Write("<hr /><hr />PORTAL_ARCHIVE_REPLY")
	if strDBType = "sqlserver" then
	strSql = "alter table [PORTAL_ARCHIVE_REPLY] drop Constraint [PK_PORTAL_ARCHIVE_REPLY];"
	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>Index dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
		
	if strDBType = "access" then	
	strSql = "DROP INDEX [PrimaryKey] on [PORTAL_ARCHIVE_REPLY] "
	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>Index dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
	
	strSql = "ALTER TABLE [PORTAL_ARCHIVE_REPLY] ADD CONSTRAINT [PrimaryKey] PRIMARY KEY([ID]);"
	my_Conn.execute(strSql)
	if err.number <> 0 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>PORTAL_ARCHIVE_REPLY Primary key added</b><br /><br />" & vbNewLine)
	end if
	err.clear
		
':::::::::::::::::::::::::::::::::: ALTER PORTAL_ARCHIVE_TOPICS TABLE :::::::::::::::::::::::::::::::::::::::::
	response.Write("<hr /><hr />PORTAL_ARCHIVE_TOPICS")
	if strDBType = "sqlserver" then
	strSql = "alter table [PORTAL_ARCHIVE_TOPICS] drop Constraint [PK_PORTAL_ARCHIVE_TOPICS];"
	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>Index dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
	
	if strDBType = "access" then	
	strSql = "DROP INDEX [PrimaryKey] ON [PORTAL_ARCHIVE_TOPICS];"
	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>Index dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
	
	strSql = "ALTER TABLE [PORTAL_ARCHIVE_TOPICS] ADD CONSTRAINT [PA_PrimaryKey] PRIMARY KEY([ID]);"
	my_Conn.execute(strSql)
	if err.number <> 0 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>PORTAL_ARCHIVE_TOPICS Primary key added</b><br /><br />" & vbNewLine)
	end if
	err.clear
		
':::::::::::::::::::::::::::::::::: ALTER PORTAL_AVATAR TABLE :::::::::::::::::::::::::::::::::::::::::
	response.Write("<hr /><hr />PORTAL_AVATAR")
	if strDBType = "sqlserver" then
	strSql = "alter table [portal_avatar] drop Constraint [PK_Portal_Avatar];"

	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>Index dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
	
    if strDBType = "access" then	
	sSQL = "DROP INDEX [PrimaryKey] ON [PORTAL_AVATAR]"
	my_Conn.execute(sSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>Index dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
	
	strSql = "ALTER TABLE [PORTAL_AVATAR] ADD CONSTRAINT [AV_PrimaryKey] PRIMARY KEY ([A_ID]);"
	my_Conn.execute(strSql)
	if err.number <> 0 then
		'Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		'ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>PORTAL_AVATAR Primary key added</b><br /><br />" & vbNewLine)
	end if
	
	':: delete data from the avatar table
	strSql = "DELETE FROM PORTAL_AVATAR WHERE A_ID <> 0;"
	my_Conn.execute(strSql)
	if err.number <> 0 then
	
	else
		Response.Write("    <b>PORTAL_AVATAR data removed</b><br /><br />" & vbNewLine)
	end if
	err.clear

	':: insert default data to the avatar table
	':: they will need to run the avatar sync.
	redim arrData(2)
	arrData(0) = "" & strTablePrefix & "AVATAR"
	arrData(1) = "A_NAME, A_URL"
	arrData(2) = "'noavatar', 'files/avatars/noavatar.gif'"
	populateB(arrData)
	Response.Write("    <b>PORTAL_AVATAR default data added</b><br />" & vbNewLine)
	Response.Write("    <b>Run the ""Avatar Sync"" from your admin area</b><br /><br />" & vbNewLine)

':::::::::::::::::::::::::::::::::: ALTER PORTAL_AVATAR2  TABLE :::::::::::::::::::::::::::::::::::::::::
	response.Write("<hr /><hr />PORTAL_AVATAR2")
	if strDBType = "sqlserver" then
	strSql = "alter table [portal_avatar2] drop Constraint [PK_Portal_Avatar2]"
	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("  <br /><b>SQL Primary Key dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	
	strSql = "if exists(SELECT [A_KEY] FROM [PORTAL_AVATAR2])ALTER TABLE [PORTAL_AVATAR2] DROP COLUMN [A_KEY];"
	my_Conn.execute(strSql)
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>A_Key field removed</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
	
	
	strSql = "PORTAL_AVATAR2,[ID] int IDENTITY CONSTRAINT [AV2_ID] PRIMARY KEY"
	alterTable2(checkIt(strSql))
	err.clear
	
	response.Write("Table PORTAL_AVATAR2 primary key added<br />")
		
		
':::::::::::::::::::::::::::::::::: ALTER PORTAL_EVENTS TABLE :::::::::::::::::::::::::::::::::::::::::
	response.Write("<hr /><hr />PORTAL_EVENTS")
	if strDBType = "sqlserver" then
	strSql = "alter table [PORTAL_EVENTS] drop Constraint [PK_PORTAL_EVENTS]"
	my_Conn.execute(strSql)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>SQL Primary Key dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	end if
	
	if strDBType = "access" then	
	sSQL = "DROP INDEX [PrimaryKey] ON [PORTAL_EVENTS]"
	my_Conn.execute(sSQL)
	if err.number <> 0 then
	  if err.number <> -2147217887 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	  end if
	else
		Response.Write("    <b>PORTAL_EVENTS Index dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	
	response.Write("<hr />")
	end if
	
	strSql="ALTER TABLE [PORTAL_EVENTS] ADD CONSTRAINT [EV_PrimaryKey] PRIMARY KEY ([EVENT_ID]);"
	my_Conn.execute(strSql)
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>PORTAL_EVENTS Primary key added</b><br /><br />" & vbNewLine)
	end if
	err.clear

':::::::::::::::::::::::::::::::::: ALTER PIC_CATEGORIES  TABLE :::::::::::::::::::::::::::::::::::::::::
	response.Write("<hr /><hr />PIC_CATEGORIES")
	strSql = "ALTER TABLE PIC_CATEGORIES ADD CAT_IMAGE TEXT(100)"
	doSQL2 checkIt(strSql),0
	err.clear
	
	response.Write("Table PIC_CATEGORIES altered<br />")

':::::::::::::::::::::::::::::::::: ALTER PIC_SUBCATEGORIES TABLE :::::::::::::::::::::::::::::::::::::::::
  response.Write("<hr /><hr />PIC_SUBCATEGORIES")
	strSql = "PIC_SUBCATEGORIES,SUBCAT_IMAGE TEXT(100),SUBCAT_THUMB TEXT(100)"
	alterTable2(checkIt(strSql))
	
	response.Write("Table PIC_SUBCATEGORIES altered<br />")

':::::::::::::::::::::::::::::::::: ALTER PIC TABLE :::::::::::::::::::::::::::::::::::::::::
	if strDBType = "access" then
	strSQL="DROP INDEX [LTURL] on [PIC]"
	else	
	strSql ="DROP INDEX [PIC].[LTURL]"
	end if
	doSQL2 checkIt(strSql),1
	
	if strDBType = "access" then
	strSQL="DROP INDEX [LURL] on [PIC]"
	else	
	strSql ="DROP INDEX [PIC].[LURL]"
	end if
	doSQL2 checkIt(strSql),1
	
	strSql = "ALTER TABLE [PIC] DROP COLUMN [LTURL]"
	doSQL2 checkIt(strSql),1

	strSql = "ALTER TABLE [PIC] DROP COLUMN [LURL]"
	doSQL2 checkIt(strSql),1

	'response.Write("Table PIC altered<br />")

':::::::::::::::::::::::::::::::::: ALTER PORTAL_EVENTS TABLE :::::::::::::::::::::::::::::::::::::::::
	strSql = "PORTAL_EVENTS,SERIES LONG NOT NULL DEFAULT 0"
	alterTable2(checkIt(strSql))
	
	response.Write("Table PORTAL_EVENTS altered<br />")
	
	strSql = "UPDATE PORTAL_EVENTS SET SERIES = 0"
	doSQL2 strSql,1

':::::::::::::::::::::::::::::::::: ALTER PORTAL_MEMBERS TABLE :::::::::::::::::::::::::::::::::::::::::
	strSql = "PORTAL_MEMBERS,M_ZIP TEXT(20) NULL,M_MSN TEXT(150) NULL,M_GLOW TEXT(50) NULL, THEME_ID VARCHAR(50) NULL, M_SHOW_BIRTHDAY LONG"
	alterTable2(checkIt(strSql))
	if strDBType = "access" then
	strSQL="DROP INDEX [THEME_ID] on [PORTAL_MEMBERS]"
	else	
	strSql ="DROP INDEX [PORTAL_MEMBERS].[THEME_ID]"
	end if
	doSQL2 checkIt(strSql),1

	if strDBType = "access" then
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN THEME_ID TEXT(50) NULL"
	else
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN THEME_ID TEXT(50)"
	end if
	doSQL2 checkIt(strSql),1


	if strDBType = "access" then
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN M_SHOW_BIRTHDAY LONG NULL DEFAULT 0"
	else
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN M_SHOW_BIRTHDAY LONG"
	end if
	doSQL2 checkIt(strSql),1
	
	if strDBType = "sqlserver" then
	on error resume next
	strSql = "ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [NN_SHOW_BIRTHDAY] CHECK([M_SHOW_BIRTHDAY] is NOT NULL);ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [DF_SHOW_BIRTHDAY] DEFAULT '0' FOR [M_SHOW_BIRTHDAY];"
	my_Conn.execute(checkIt(strSql))
    if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>M_SHOW_BIRTHDAY default 0 changed</b><br /><br />" & vbNewLine)
	end if
	err.clear	
	on error goto 0	
	end if
	
	if strDBType = "sqlserver" then
	strSql = "ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [NN_THEME_ID] CHECK([THEME_ID] is NOT NULL);ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [DF_THEME_ID] DEFAULT '0' FOR [THEME_ID];"
	on error resume next
	my_Conn.execute(checkIt(strSql))
    if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>THEME_ID field Not Null changed</b><br /><br />" & vbNewLine)
	end if
	err.clear
	on error goto 0		
	end if

	redim indexes(2)
	indexes(0) = "CREATE INDEX [THEME_ID] ON [PORTAL_MEMBERS]([THEME_ID]);"
	indexes(1) = "CREATE INDEX [M_SHOW_BIRTHDAY] ON [PORTAL_MEMBERS]([M_SHOW_BIRTHDAY]);"
	indexes(2) = "CREATE INDEX [M_GLOW] ON [PORTAL_MEMBERS]([M_GLOW]);"
	createIndx(indexes)
		
	strSql = "UPDATE PORTAL_MEMBERS SET THEME_ID = '0'"
	doSQL2 strSql,1
	err.clear
	
	strSql = "UPDATE PORTAL_MEMBERS SET M_SHOW_BIRTHDAY = 0"
	doSQL2 strSql,1
	
	response.Write("Table PORTAL_MEMBERS altered<br />")

':::::::::::::::::::::::::::::::::: ALTER PORTAL_MEMBERS_PENDING TABLE :::::::::::::::::::::::::::::::::::::::::
	strSql = "PORTAL_MEMBERS_PENDING,M_ZIP TEXT(20),M_MSN TEXT(150),THEME_ID VARCHAR(50) NULL,M_GLOW TEXT(50) NULL"
	alterTable2(checkIt(strSql))
	
	response.Write("Table PORTAL_MEMBERS_PENDING altered<br />")

':::::::::::::::::::::::::::::::::: UPDATE PORTAL_MODS  TABLE :::::::::::::::::::::::::::::::::::::::::
  if strDBType = "sqlserver" then
	strSql = "alter table [PORTAL_MODS] drop Constraint [PK_PORTAL_MODS]"
	on error resume next
	my_Conn.execute(strSql)
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>SQL Primary Key dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	response.Write("<hr />")
	
	strSql = "if exists(SELECT [m_ID] FROM [PORTAL_MODS])ALTER TABLE [PORTAL_MODS] DROP COLUMN [m_ID];"
	my_Conn.execute(strSql)
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>A_Key field removed</b><br /><br />" & vbNewLine)
	end if
	err.clear
	on error goto 0
	response.Write("<hr />")
  end if
	
	strSql = "PORTAL_MODS,[ID] int IDENTITY CONSTRAINT [MOD_ID] PRIMARY KEY"
	alterTable2(checkIt(strSql))

	sSql = "SELECT M_NAME FROM PORTAL_MODS WHERE M_CODE = 'slColumns'"
	set rsMod = my_Conn.execute(sSql)
	if rsMod.eof then
	  strSql = "INSERT INTO PORTAL_MODS (M_CODE, M_VALUE, M_NAME) values ('slColumns','1','news')"
	  doSQL2 strSql,1
	  response.Write("Table PORTAL_MODS updated<br />")
	end if
	set rsMod = nothing
	
	sSql = "SELECT M_NAME FROM PORTAL_MODS WHERE M_CODE = 'slDefimg'"
	set rsMod = my_Conn.execute(sSql)
	if rsMod.eof then
	  strSql = "INSERT INTO PORTAL_MODS (M_CODE, M_VALUE, M_NAME) values ('slDefimg','images/news.gif','news')"
	  doSQL2 strSql,1
	  response.Write("Table PORTAL_MODS updated<br />")
	end if
	


	'::::::::::::::::::: ALTER PORTAL_TOPICS TABLE :::::::::::::::::::::::::::
response.Write("<br />:::::::::::::::: ALTER PORTAL_TOPICS TABLE  ::::::::::::::<br />")
		
strSql = "CREATE INDEX [T_AUTHOR] ON [PORTAL_TOPICS]([T_AUTHOR]);"
doSQL2 strSql,0
strSql= "CREATE INDEX [T_DATE] ON [PORTAL_TOPICS]([T_DATE]);"
doSQL2 strSql,0
strSql="CREATE INDEX [T_STATUS] ON [PORTAL_TOPICS]([T_STATUS]);"
doSQL2 strSql,0
strSql="CREATE INDEX [T_LAST_POST_AUTHOR] ON [PORTAL_TOPICS]([T_LAST_POST_AUTHOR]);"
doSQL2 strSql,0
strSql="CREATE INDEX [T_LAST_POST] ON [PORTAL_TOPICS]([T_LAST_POST]);"
doSQL2 strSql,0
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b> 4 indexes added</b><br /><br />" & vbNewLine)
	end if
	err.clear		
	
err.clear
		':::::::::::::::::::::::::::::::::: ALTER PORTAL_REPLY TABLE :::::::::::::::::::::::::::::::::::::::::
		
strSql = "CREATE INDEX [R_DATE] ON [PORTAL_REPLY]([R_DATE]);"
doSQL2 strSql,0
strSql="CREATE INDEX [R_AUTHOR] ON [PORTAL_REPLY]([R_AUTHOR]);"
doSQL2 strSql,0
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b> 4 indexes added</b><br /><br />" & vbNewLine)
	end if
	err.clear		
	response.Write("PORTAL_REPLY Updated for 2 Indexes")	
	
		':::::::::::::::::::::::::::::::::: ALTER PORTAL_Online TABLE :::::::::::::::::::::::::::::::::::::::::
  response.write "PORTAL_Online"  
	strSql = "ALTER TABLE [PORTAL_ONLINE] ADD [ID] int IDENTITY CONSTRAINT [PO_ID] PRIMARY KEY"
	doSQL2 checkIt(strSql),0
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>Table PORTAL_ONLINE primary key: ID added</b><br /><br />" & vbNewLine)
	end if
	err.clear
	strSql = "ALTER TABLE PORTAL_ONLINE ALTER COLUMN M_BROWSE MEMO"
	on error resume next
	my_Conn.execute(checkIt(strSql))
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>Table PORTAL_ONLINE M_BROWSE UPDATED</b><br /><br />" & vbNewLine)
	end if
	err.clear
	
	if strDBType = "sqlserver" then
	strSql = "alter table [PORTAL_MODERATOR] drop Constraint [DF__PORTAL_MO__MOD_T__1A9EF37A];"
		my_Conn.execute(strSql)
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>Defalt Constraint dropped</b><br /><br />" & vbNewLine)
	end if
	err.clear
	end if
		
	strSql = "ALTER TABLE [PORTAL_MODERATOR] ALTER COLUMN [MOD_TYPE] LONG"
	my_Conn.execute(checkIt(strSql))
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>Table PORTAL_MODERATORS UPDATED</b><br /><br />" & vbNewLine)
	end if
	err.clear
'::::::::::::::::::::::::::::UPDATE Members with no avatar:::::::::::::::::::::::::::;::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	strSQL="UPDATE PORTAL_MEMBERS SET [M_AVATAR_URL] = 'files/avatars/noavatar.gif' WHERE MEMBER_ID <> 0"
		my_Conn.execute(checkIt(strSql))
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>Table PORTAL_MEMBERS - all member avatars RESET to default</b><br /><br />" & vbNewLine)
	end if
	
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	response.Write "<br /><br /><br /><br />" & dbHits & " database hits<br /><br />"
	err.clear
	on error goto 0
end sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'############################################################################################
'############################################################################################
'############################################################################################
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub update15b3()
	':: UPGRADES FROM v15b3 to v2.0
	':::::::::::::::: ALTER PORTAL_MEMBERS_PENDING TABLE ::::::::::::::::::::::::::::::::::::
	strSql = "PORTAL_MEMBERS_PENDING,THEME_ID VARCHAR(50) NULL"
	alterTable2(checkIt(strSql))
	
	response.Write("Table PORTAL_MEMBERS_PENDING altered<br />")
	if strDBType = "access" then
	strSQL="DROP INDEX [M_GLOW] on [PORTAL_MEMBERS]"
	else	
	strSql ="DROP INDEX [PORTAL_MEMBERS].[M_GLOW]"
	end if
	on error resume next
	doSQL2 checkIt(strSql),1
	err.clear
	
	':::::::::::::::: ALTER PORTAL_MEMBERS TABLE ::::::::::::::::::::::::::::::::::::
	
	'strSql = "PORTAL_MEMBERS, THEME_ID TEXT(50) NULL, M_GLOW TEXT(50) NULL, M_SHOW_BIRTHDAY LONG NULL"
	strSql = "PORTAL_MEMBERS, THEME_ID TEXT(50) NULL, M_SHOW_BIRTHDAY LONG NULL"
	alterTable2(checkIt(strSql))
	strSql = "PORTAL_CONFIG, C_STRLOCKDOWN BYTE NULL, C_MODULES TEXT(50) NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE PORTAL_CONFIG SET C_STRLOCKDOWN = 0"
	on error resume next
	doSQL2 checkIt(strSql),0
	err.clear
	
	if strDBType = "access" then
	strSQL="DROP INDEX [THEME_ID] on [PORTAL_MEMBERS]"
	else	
	strSql ="DROP INDEX [PORTAL_MEMBERS].[THEME_ID]"
	end if
	on error resume next
	doSQL2 checkIt(strSql),1
	err.clear
	if strDBType = "access" then
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN THEME_ID TEXT(50) NULL DEFAULT 0"
	else
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN THEME_ID TEXT(50)"
	end if
	on error resume next
	doSQL2 checkIt(strSql),1
	err.clear
	if strDBType = "access" then
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN M_SHOW_BIRTHDAY LONG NULL DEFAULT 0"
	else
	strSql = "ALTER TABLE PORTAL_MEMBERS ALTER COLUMN M_SHOW_BIRTHDAY LONG"
	end if
	on error resume next
	doSQL2 checkIt(strSql),1
	
	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>THEME_ID field type changed</b><br /><br />" & vbNewLine)
	end if
	err.clear
	
	if strDBType = "sqlserver" then
	strSql = "ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [NN_SHOW_BIRTHDAY] CHECK([M_SHOW_BIRTHDAY] is NOT NULL);ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [DF_SHOW_BIRTHDAY] DEFAULT '0' FOR [M_SHOW_BIRTHDAY];"
	my_Conn.execute(checkIt(strSql))
    if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>M_SHOW_BIRTHDAY default 0 changed</b><br /><br />" & vbNewLine)
	end if
	err.clear		
	end if
	
	if strDBType = "sqlserver" then
	strSql = "ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [NN_THEME_ID] CHECK([THEME_ID]is NOT NULL);ALTER TABLE [PORTAL_MEMBERS] ADD CONSTRAINT [DF_THEME_ID]DEFAULT '0' FOR [THEME_ID];"
	my_Conn.execute(checkIt(strSql))
    if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <b>THEME_ID field Not Null changed</b><br /><br />" & vbNewLine)
	end if
	err.clear		
	end if

	redim indexes(1)
	indexes(0) = "CREATE INDEX [THEME_ID] ON [PORTAL_MEMBERS]([THEME_ID]);"
	indexes(1) = "CREATE INDEX [M_SHOW_BIRTHDAY] ON [PORTAL_MEMBERS]([M_SHOW_BIRTHDAY]);"
	createIndx(indexes)
		
	strSql = "UPDATE PORTAL_MEMBERS SET THEME_ID = '0' WHERE MEMBER_ID > 0"
	my_Conn.execute(strSql)
	err.clear
	
	strSql = "UPDATE PORTAL_MEMBERS SET M_SHOW_BIRTHDAY = 0 WHERE MEMBER_ID > 0"
	doSQL2 strSql,1
	
	response.Write("Table PORTAL_MEMBERS altered<br /><hr />")
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	if strDBType = "access" then
	strSQL="DROP INDEX [LTURL] on [PIC]"
	else	
	strSql ="DROP INDEX [PIC].[LTURL]"
	end if
	on error resume next
	doSQL2 checkIt(strSql),1
	err.clear
	if strDBType = "access" then
	strSQL="DROP INDEX [LURL] on [PIC]"
	else	
	strSql ="DROP INDEX [PIC].[LURL]"
	end if
	on error resume next
	doSQL2 checkIt(strSql),1
	err.clear

end sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'############################################################################################
'############################################################################################
'############################################################################################
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'%%%%%%%%%% UPDATE FROM v2.0 to v2.1 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
sub update20x21()
':::::::::: ALTER PORTAL_CONFIG TABLE add field ::::::::::::::::
	response.Write("<hr /><hr />")

	strSql = "PORTAL_CONFIG,C_SECIMAGE INTEGER"
	alterTable2(checkIt(strSql))

	'-------------------- populate table with default values --------------------------
	strSql = "UPDATE PORTAL_CONFIG SET C_SECIMAGE = 0 WHERE CONFIG_ID = 1"
	doSQL2 checkIt(strSql),1
	
	Response.Write("<b>Table PORTAL_CONFIG - updated for v2.0 - v2.1 fields and data</b><br /><br />" & vbNewLine)
end sub

%>