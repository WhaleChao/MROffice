/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Insert > Icons catalog every editor's gallery reads: several hundred
 * line icons (lucide, ISC licence -- bundling and redistribution permitted),
 * each in a category, with Chinese and English search words. Data only; the
 * drawings themselves are in `./glyphs`, loaded when a gallery opens.
 *
 * PowerPoint's Insert > Icons (Microsoft Support, "Insert icons in Microsoft
 * 365"): categories to jump between, a search box, several icons chosen and
 * inserted together, each an SVG graphic its Fill recolours. This catalog
 * supplies the first two; the gallery and the host do the rest.
 */

export const OFFICE_ICON_CATEGORIES = Object.freeze([
  "people",
  "business",
  "communication",
  "education",
  "arrows",
  "interface",
  "technology",
  "data",
  "files",
  "media",
  "nature",
  "weather",
  "transport",
  "travel",
  "food",
  "medical",
  "sports",
  "celebration",
  "home",
  "security",
  "tools",
  "shopping",
  "time",
  "symbols",
] as const);

export type OfficeIconCategory = (typeof OFFICE_ICON_CATEGORIES)[number];

export interface OfficeIconEntry {
  /** Stable id: the gallery's test id and the inserted file's name. */
  readonly id: string;
  /** The lucide component name in `./glyphs`. */
  readonly glyph: string;
  readonly categories: readonly OfficeIconCategory[];
  readonly zh: string;
  readonly en: string;
  /** Lower-case Chinese and English words a search matches. */
  readonly keywords: string;
}

/** The ids the gallery shipped with before the catalog, kept for them, and
 * the two glyphs whose own names those ids took. */
const LEGACY_IDS: Readonly<Record<string, string>> = {
  Sparkles: "sparkles", Check: "check", Camera: "camera", Image: "image",
  MessageSquare: "comment", MonitorPlay: "presentation", ChartNoAxesCombined: "chart",
  Table2: "table", Palette: "palette", Timer: "timer", Volume2: "audio", Film: "video",
  // Their own names are taken by two of the ids above.
  Presentation: "presentation-board", Video: "video-call",
};

// category|LucideName|中文名（、分隔）|english keywords
const RAW = `
people|User|人、用户、个人|person profile account
people|Users|多人、团队、用户组|group team people
people|UserPlus|添加用户、新成员|add user invite
people|UserCheck|已验证用户、确认人员|verified approved
people|UserX|移除用户|remove delete
people|UserRound|头像、人像|avatar portrait
people|UsersRound|群组、团体|group community
people|Contact|联系人、通讯录|contact address book
people|Baby|婴儿、儿童|baby child infant
people|PersonStanding|站立的人、无障碍|person standing accessibility
people|Accessibility|无障碍、辅助功能|accessibility disabled
people|Hand|手、举手|hand raise stop
people|HandHeart|关爱、慈善|care charity
people|Handshake|握手、合作|handshake deal partnership
people|ThumbsUp|点赞、赞同|like approve good
people|ThumbsDown|不赞同、差评|dislike bad
people|Smile|微笑、开心|happy smile face
people|Frown|难过、不开心|sad face
people|Laugh|大笑|laugh joy
people|Meh|一般、无表情|neutral face
people|Angry|生气|angry mad
people|Brain|大脑、思考|brain mind think
people|Footprints|脚印、足迹|footprints steps
people|Ear|耳朵、听|ear hear listen
people|Eye|眼睛、查看|eye view see
people|Speech|演讲、说话|speech talk
people|Fingerprint|指纹、身份|fingerprint identity biometrics
business|Briefcase|公文包、工作|briefcase work job
business|Building|大楼、公司|building office company
business|Building2|办公楼、企业|office building enterprise
business|Factory|工厂、制造|factory manufacturing
business|Landmark|银行、机构|bank institution landmark
business|Store|商店、门店|store shop retail
business|Presentation|演示、汇报|presentation slides meeting
business|Target|目标、靶心|target goal aim
business|Trophy|奖杯、成就|trophy award win
business|Award|奖章、荣誉|award badge medal
business|Medal|奖牌|medal prize
business|Crown|皇冠、领导|crown leader king
business|Rocket|火箭、启动、增长|rocket launch startup
business|Lightbulb|灯泡、创意、想法|idea lightbulb innovation
business|Puzzle|拼图、解决方案|puzzle solution
business|Scale|天平、公平、法律|scale balance law justice
business|Gavel|法槌、法律|gavel law judge
business|Signature|签名|signature sign
business|Stamp|印章|stamp approve
business|ClipboardList|清单、任务|checklist clipboard tasks
business|ClipboardCheck|核对清单、完成|clipboard done
business|Kanban|看板|kanban board
business|Workflow|工作流程|workflow process
business|Network|网络、组织结构|network organization structure
business|Waypoints|路径、节点|waypoints path
business|Goal|目标旗|goal flag
business|Flag|旗帜、标记|flag mark milestone
business|Milestone|里程碑|milestone signpost
business|BadgeCheck|认证、徽章|verified badge
business|Megaphone|喇叭、宣传、公告|megaphone announce marketing
business|Mic|麦克风|microphone speak
communication|Mail|邮件、信封|mail email envelope
communication|MailOpen|已读邮件|mail open read
communication|Send|发送|send paper plane
communication|Inbox|收件箱|inbox
communication|MessageSquare|消息、批注|message comment chat
communication|MessageCircle|聊天、对话|chat bubble conversation
communication|MessagesSquare|多条消息、讨论|discussion messages
communication|Phone|电话|phone call
communication|PhoneCall|来电、通话|phone call ringing
communication|Smartphone|手机|smartphone mobile
communication|AtSign|艾特、邮箱符号|at sign email
communication|Bell|铃铛、通知|bell notification alert
communication|BellRing|提醒|bell ringing reminder
communication|Rss|订阅|rss feed subscribe
communication|Share2|分享|share
communication|Radio|广播、收音机|radio broadcast
communication|Podcast|播客|podcast
communication|Video|视频通话|video camera call
communication|Voicemail|语音信箱|voicemail
communication|Languages|语言、翻译|language translate
communication|Globe|地球、全球、网站|globe world web
communication|Satellite|卫星|satellite
communication|Antenna|天线、信号|antenna signal
communication|Newspaper|报纸、新闻|newspaper news
education|GraduationCap|学士帽、毕业|graduation education school
education|School|学校|school
education|BookOpen|打开的书、阅读|book read open
education|Book|书|book
education|Library|图书馆、书架|library books
education|Notebook|笔记本|notebook
education|NotebookPen|写笔记|notes write
education|Pencil|铅笔|pencil write edit
education|PenTool|钢笔、设计|pen tool design
education|Highlighter|荧光笔|highlighter
education|Ruler|直尺|ruler measure
education|Calculator|计算器|calculator math
education|FlaskConical|烧瓶、实验|flask lab science
education|Microscope|显微镜|microscope science
education|Atom|原子、科学|atom science physics
education|Dna|DNA、基因|dna gene biology
education|Telescope|望远镜|telescope astronomy
education|Backpack|书包|backpack school bag
education|Sigma|求和、数学|sigma sum math
education|Pi|圆周率|pi math
education|Shapes|图形、几何|shapes geometry
education|Lectern|讲台|lectern podium
arrows|ArrowUp|向上箭头|arrow up
arrows|ArrowDown|向下箭头|arrow down
arrows|ArrowLeft|向左箭头|arrow left
arrows|ArrowRight|向右箭头|arrow right
arrows|ArrowUpRight|右上箭头、增长|arrow up right growth
arrows|ArrowDownRight|右下箭头、下降|arrow down right decline
arrows|ArrowBigUp|粗箭头向上|big arrow up
arrows|ArrowBigRight|粗箭头向右|big arrow right
arrows|ArrowLeftRight|双向箭头、交换|swap exchange two way
arrows|ArrowUpDown|上下箭头、排序|sort up down
arrows|ChevronRight|右尖括号|chevron right next
arrows|ChevronsRight|快进箭头|chevrons forward
arrows|CornerDownRight|转折箭头|corner arrow
arrows|Undo2|撤销、返回|undo back
arrows|Redo2|重做、前进|redo forward
arrows|RefreshCw|刷新、循环|refresh cycle sync
arrows|RotateCw|顺时针旋转|rotate clockwise
arrows|Repeat|重复|repeat loop
arrows|Shuffle|随机、打乱|shuffle random
arrows|Move|移动|move arrows
arrows|Maximize2|放大、全屏|expand maximize
arrows|Minimize2|缩小|collapse minimize
arrows|TrendingUp|上升趋势|trend up growth
arrows|TrendingDown|下降趋势|trend down decline
arrows|Split|分叉|split branch
arrows|Merge|合并|merge
arrows|CircleArrowRight|圆形右箭头|circle arrow right
arrows|MoveRight|长箭头向右|long arrow right
interface|Check|对勾、完成|check done tick
interface|CheckCheck|双对勾、已读|double check read
interface|X|关闭、错误|close cross error
interface|Plus|加号、添加|plus add
interface|Minus|减号|minus remove
interface|CircleCheck|圆形对勾|check circle success
interface|CircleX|圆形叉号|x circle error
interface|CircleAlert|警告|alert warning circle
interface|TriangleAlert|三角警告|warning triangle caution
interface|Info|信息|info information
interface|CircleHelp|帮助、问号|help question
interface|Search|搜索、放大镜|search magnifier
interface|Settings|设置、齿轮|settings gear
interface|SlidersHorizontal|调节、筛选|adjust filter sliders
interface|Filter|筛选、漏斗|filter funnel
interface|Menu|菜单|menu hamburger
interface|House|主页、房子|home house
interface|LayoutGrid|网格布局|grid layout
interface|LayoutDashboard|仪表板|dashboard
interface|List|列表|list
interface|ListChecks|任务清单|checklist tasks
interface|Link|链接|link chain
interface|Paperclip|回形针、附件|attachment paperclip
interface|Bookmark|书签|bookmark save
interface|Tag|标签|tag label
interface|Pin|图钉|pin
interface|Download|下载|download
interface|Upload|上传|upload
interface|Trash2|删除、垃圾桶|delete trash bin
interface|Copy|复制|copy duplicate
interface|Clipboard|剪贴板|clipboard paste
interface|Eye|查看、可见|view visible
interface|EyeOff|隐藏|hidden hide
interface|Lock|锁、锁定|lock locked
interface|LockOpen|解锁|unlock open
interface|Key|钥匙|key password
interface|Power|电源、开关|power on off
interface|ToggleRight|开关|toggle switch
interface|Star|星星、收藏|star favorite rating
interface|Heart|心、喜欢|heart love like
interface|Sparkles|闪光|sparkles magic ai
interface|Zap|闪电、快速|lightning fast power
interface|Grip|拖动手柄|grip drag
interface|Ellipsis|更多|more ellipsis
interface|ExternalLink|外部链接|external link
interface|LogIn|登录|login sign in
interface|LogOut|退出、登出|logout sign out
technology|Laptop|笔记本电脑|laptop computer
technology|Monitor|显示器|monitor screen display
technology|MonitorPlay|演示、播放屏幕|presentation screen play
technology|Tablet|平板电脑|tablet
technology|Keyboard|键盘|keyboard
technology|Mouse|鼠标|mouse
technology|Printer|打印机|printer print
technology|Server|服务器|server
technology|Database|数据库|database storage
technology|HardDrive|硬盘|hard drive disk
technology|Cpu|处理器、芯片|cpu chip processor
technology|MemoryStick|U盘、内存|usb memory
technology|Cloud|云|cloud
technology|CloudUpload|云上传|cloud upload
technology|CloudDownload|云下载|cloud download
technology|Wifi|无线网络|wifi wireless
technology|Bluetooth|蓝牙|bluetooth
technology|Battery|电池|battery
technology|BatteryCharging|充电|charging battery
technology|Plug|插头|plug power
technology|Code|代码|code programming
technology|Terminal|终端、命令行|terminal console
technology|Bug|漏洞、错误|bug debug
technology|Bot|机器人|bot robot ai
technology|QrCode|二维码|qr code
technology|ScanLine|扫描|scan
technology|Gamepad2|游戏手柄|gamepad game controller
technology|Headphones|耳机|headphones audio
technology|Watch|手表|watch smartwatch
technology|Webcam|网络摄像头|webcam
technology|Router|路由器|router network
technology|Usb|USB|usb
data|ChartNoAxesCombined|图表|chart data
data|ChartColumn|柱形图|bar chart column
data|ChartBar|条形图|bar chart
data|ChartLine|折线图|line chart
data|ChartPie|饼图|pie chart
data|ChartArea|面积图|area chart
data|ChartScatter|散点图|scatter chart
data|ChartCandlestick|K线图|candlestick stock chart
data|ChartNetwork|网络图|network chart
data|ChartGantt|甘特图|gantt timeline
data|Table2|表格|table grid
data|Sheet|电子表格|spreadsheet sheet
data|Activity|活动、心电图|activity pulse
data|Gauge|仪表、速度|gauge speedometer
data|Percent|百分比|percent
data|Hash|井号、编号|hash number
data|Binary|二进制|binary data
data|Layers|图层|layers
data|Funnel|漏斗|funnel conversion
files|File|文件|file document
files|FileText|文本文档|text document
files|FilePlus|新建文件|new file
files|FileCheck|已确认文件|file check approved
files|FileSpreadsheet|表格文件|spreadsheet file
files|FileImage|图片文件|image file
files|FileVideo|视频文件|video file
files|FileAudio|音频文件|audio file
files|FileArchive|压缩文件|archive zip
files|FileCode|代码文件|code file
files|Files|多个文件|files documents
files|Folder|文件夹|folder
files|FolderOpen|打开的文件夹|folder open
files|Archive|归档|archive box
files|ScrollText|卷轴、合同|scroll contract
files|NotepadText|记事本|notepad
files|StickyNote|便签|sticky note
files|Receipt|收据、发票|receipt invoice
media|Image|图片|image picture photo
media|Images|多张图片、相册|gallery images
media|Camera|摄像头、相机|camera photo
media|Film|视频、胶片|film movie video
media|Clapperboard|场记板、电影|clapperboard movie
media|Music|音乐|music note
media|Volume2|音频、音量|audio volume sound
media|VolumeX|静音|mute
media|Play|播放|play
media|Pause|暂停|pause
media|CircleStop|停止|stop
media|SkipForward|下一首|skip forward next
media|Disc3|光盘|disc cd
media|Tv|电视|tv television
media|Projector|投影仪|projector
media|Palette|调色板|palette colors art
media|Brush|画笔|brush paint
media|Paintbrush|涂刷|paintbrush
media|Aperture|光圈|aperture camera
media|Drama|戏剧、面具|drama theater masks
nature|Leaf|叶子|leaf nature eco
nature|Trees|树林|trees forest
nature|TreePine|松树|pine tree
nature|TreeDeciduous|大树|tree
nature|Flower|花|flower
nature|Flower2|花朵|blossom flower
nature|Sprout|嫩芽、成长|sprout growth plant
nature|Clover|四叶草、幸运|clover luck
nature|Mountain|山|mountain
nature|MountainSnow|雪山|snow mountain
nature|Waves|波浪、海|waves sea water
nature|Droplet|水滴|water drop
nature|Flame|火焰|fire flame
nature|Recycle|回收、环保|recycle eco
nature|Earth|地球|earth planet
nature|Bird|鸟|bird
nature|Cat|猫|cat pet
nature|Dog|狗|dog pet
nature|Fish|鱼|fish
nature|Rabbit|兔子|rabbit
nature|Turtle|乌龟|turtle
nature|Squirrel|松鼠|squirrel
nature|Bug|昆虫|insect bug
nature|PawPrint|爪印、宠物|paw pet
nature|Shell|贝壳|shell
nature|Feather|羽毛|feather
weather|Sun|太阳、晴天|sun sunny
weather|Moon|月亮、夜晚|moon night
weather|Cloud|云、阴天|cloud cloudy
weather|CloudRain|下雨|rain
weather|CloudSnow|下雪|snow
weather|CloudLightning|雷雨|storm thunder lightning
weather|CloudSun|多云|partly cloudy
weather|CloudFog|雾|fog mist
weather|Snowflake|雪花、冬天|snowflake winter cold
weather|Umbrella|雨伞|umbrella rain
weather|Wind|风|wind
weather|Thermometer|温度计|thermometer temperature
weather|Rainbow|彩虹|rainbow
weather|Sunrise|日出|sunrise morning
weather|Sunset|日落|sunset evening
weather|Tornado|龙卷风|tornado
weather|Haze|霾|haze
transport|Car|汽车|car vehicle
transport|Bus|公交车|bus
transport|Truck|卡车、货运|truck delivery
transport|TrainFront|火车|train
transport|TramFront|电车|tram
transport|Plane|飞机|plane flight
transport|Ship|轮船|ship boat
transport|Sailboat|帆船|sailboat
transport|Bike|自行车|bike bicycle
transport|Motorbike|摩托车|motorbike
transport|Rocket|火箭|rocket space
transport|Fuel|加油站|fuel gas
transport|ParkingMeter|停车|parking
transport|TrafficCone|路锥|traffic cone
transport|Ambulance|救护车|ambulance
transport|Tractor|拖拉机|tractor farm
transport|CarTaxiFront|出租车|taxi
transport|Anchor|锚|anchor
travel|MapPin|定位、地点|location pin place
travel|Map|地图|map
travel|Compass|指南针|compass navigation
travel|Navigation|导航|navigation
travel|Route|路线|route
travel|Luggage|行李|luggage suitcase
travel|Hotel|酒店|hotel
travel|Tent|帐篷、露营|tent camping
travel|TicketsPlane|机票|ticket flight
travel|Ticket|门票|ticket
travel|Signpost|路标|signpost
travel|Castle|城堡|castle
travel|Church|教堂|church
travel|Hospital|医院|hospital
travel|University|大学|university
travel|Palmtree|棕榈树、度假|palm tree vacation beach
travel|Globe2|全球|globe world
travel|Earth|世界|world earth
food|Coffee|咖啡|coffee cafe
food|CupSoda|饮料|soda drink
food|Wine|红酒|wine
food|Beer|啤酒|beer
food|Martini|鸡尾酒|cocktail
food|Pizza|披萨|pizza
food|Sandwich|三明治|sandwich
food|Soup|汤|soup
food|Salad|沙拉|salad
food|Apple|苹果|apple fruit
food|Cherry|樱桃|cherry
food|Grape|葡萄|grape
food|Citrus|柑橘|citrus orange lemon
food|Banana|香蕉|banana
food|Carrot|胡萝卜|carrot vegetable
food|Cake|蛋糕|cake birthday
food|CakeSlice|蛋糕块|cake slice
food|Cookie|饼干|cookie
food|IceCreamCone|冰淇淋|ice cream
food|Croissant|牛角包|croissant
food|Egg|鸡蛋|egg
food|Drumstick|鸡腿|chicken drumstick
food|Utensils|餐具|utensils dining
food|ChefHat|厨师帽|chef cooking
food|CookingPot|锅|cooking pot
food|Milk|牛奶|milk
food|Popcorn|爆米花|popcorn
medical|Stethoscope|听诊器|stethoscope doctor
medical|HeartPulse|心跳|heart pulse health
medical|Pill|药丸|pill medicine
medical|Syringe|注射器|syringe vaccine
medical|Thermometer|体温计|thermometer fever
medical|Bandage|绷带|bandage
medical|Cross|十字|cross medical
medical|Hospital|医院|hospital
medical|Ambulance|急救|ambulance emergency
medical|Activity|生命体征|vital signs
medical|Brain|脑科|brain neurology
medical|Bone|骨头|bone
medical|Dna|基因|dna genetics
medical|Microscope|化验|lab test
medical|ShieldPlus|健康保障|health insurance
medical|Accessibility|康复|rehabilitation
medical|Tablets|药片|tablets medicine
medical|Glasses|眼镜|glasses
sports|Dumbbell|哑铃、健身|dumbbell fitness gym
sports|Bike|骑行|cycling
sports|Footprints|跑步|running
sports|Trophy|冠军|champion trophy
sports|Medal|奖牌|medal
sports|Volleyball|排球|volleyball
sports|Goal|球门|goal
sports|Timer|计时器|timer stopwatch
sports|Mountain|登山|climbing
sports|Waves|游泳|swimming
sports|Target|射箭|archery target
sports|Flag|终点|finish flag
sports|Gamepad2|电竞|esports gaming
sports|Dices|骰子|dice games
sports|Puzzle|益智|puzzle
celebration|PartyPopper|派对、庆祝|party celebration
celebration|Gift|礼物|gift present
celebration|Cake|生日蛋糕|birthday cake
celebration|Balloon|气球|balloon
celebration|Sparkles|闪耀|sparkle
celebration|Star|星星|star
celebration|Heart|爱心|love heart
celebration|Crown|皇冠|crown
celebration|Wine|干杯|cheers toast
celebration|Music|音乐|music
celebration|Candy|糖果|candy
celebration|TreePine|圣诞树|christmas tree
celebration|Snowflake|冬季节日|winter holiday
celebration|Flower|鲜花|flowers
celebration|Ribbon|丝带|ribbon
celebration|Gem|宝石|gem diamond
home|House|房子|house home
home|Sofa|沙发|sofa couch
home|Bed|床|bed sleep
home|Lamp|台灯|lamp light
home|DoorOpen|门|door
home|Bath|浴缸|bath
home|ShowerHead|淋浴|shower
home|WashingMachine|洗衣机|washing machine laundry
home|Refrigerator|冰箱|fridge
home|Microwave|微波炉|microwave
home|Armchair|扶手椅|armchair
home|Warehouse|仓库|warehouse
home|Fence|栅栏|fence
home|Shirt|衬衫、衣服|shirt clothes
home|Key|钥匙|key
home|Lightbulb|灯泡|light bulb
security|Shield|盾牌、安全|shield security protection
security|ShieldCheck|安全认证|secure verified
security|ShieldAlert|安全警告|security alert
security|Lock|加密、锁|lock secure
security|KeyRound|密钥|key
security|Fingerprint|指纹|fingerprint
security|ScanFace|人脸识别|face id scan
security|Siren|警报|siren alarm
security|Cctv|监控摄像头|cctv surveillance
security|Ban|禁止|ban forbidden
security|CircleOff|关闭|off disabled
security|Vault|保险库|vault safe
tools|Wrench|扳手|wrench tool
tools|Hammer|锤子|hammer
tools|Scissors|剪刀|scissors cut
tools|PaintBucket|油漆桶|paint bucket
tools|Pipette|吸管|eyedropper pipette
tools|Ruler|尺子|ruler
tools|Construction|施工|construction
tools|HardHat|安全帽|hard hat
tools|Drill|电钻|drill
tools|Axe|斧头|axe
tools|Shovel|铲子|shovel
tools|Magnet|磁铁|magnet
tools|Cog|齿轮|cog gear
tools|Plug|插头|plug
tools|Flashlight|手电筒|flashlight torch
tools|Telescope|望远镜|telescope
shopping|ShoppingCart|购物车|shopping cart
shopping|ShoppingBag|购物袋|shopping bag
shopping|CreditCard|信用卡|credit card payment
shopping|Wallet|钱包|wallet
shopping|Banknote|纸币|money cash
shopping|Coins|硬币|coins money
shopping|PiggyBank|存钱罐|savings piggy bank
shopping|DollarSign|美元|dollar
shopping|Euro|欧元|euro
shopping|JapaneseYen|人民币、日元|yuan yen currency
shopping|PoundSterling|英镑|pound
shopping|Bitcoin|比特币|bitcoin crypto
shopping|BadgePercent|折扣|discount sale
shopping|Package|包裹|package parcel
shopping|Barcode|条形码|barcode
shopping|Store|店铺|shop
shopping|Gift|礼品|gift
shopping|Tags|价签|tags price
shopping|HandCoins|付款|payment hand coins
shopping|Landmark|银行|bank
time|Clock|时钟、时间|clock time
time|AlarmClock|闹钟|alarm clock
time|Timer|计时器|timer
time|Hourglass|沙漏|hourglass
time|Calendar|日历|calendar date
time|CalendarDays|日程|schedule calendar
time|CalendarCheck|已安排|scheduled
time|History|历史|history
time|Watch|手表|watch
time|CalendarClock|截止时间|deadline
symbols|Circle|圆形|circle
symbols|Square|方形|square
symbols|Triangle|三角形|triangle
symbols|Hexagon|六边形|hexagon
symbols|Pentagon|五边形|pentagon
symbols|Octagon|八边形|octagon
symbols|Diamond|菱形|diamond
symbols|Star|五角星|star
symbols|Heart|心形|heart
symbols|Infinity|无限|infinity
symbols|Asterisk|星号|asterisk
symbols|Hash|井号|hash
symbols|Quote|引号|quote
symbols|Ampersand|与号|ampersand and
symbols|CircleDollarSign|货币符号|currency
symbols|Copyright|版权|copyright
symbols|Equal|等号|equal
symbols|Divide|除号|divide
symbols|Radiation|辐射|radiation
symbols|Biohazard|生物危害|biohazard
symbols|Command|命令键|command
symbols|Option|选项键|option
symbols|Badge|徽章|badge
symbols|Component|组件|component
symbols|Spline|曲线|curve spline
symbols|Squircle|圆角方形|squircle
tools|PaintRoller|滚筒刷、粉刷|paint roller
`;

function humanize(glyph: string): string {
  const words = glyph.replace(/\d+$/, "").replace(/([a-z])([A-Z])/g, "$1 $2").toLowerCase();
  return words.charAt(0).toUpperCase() + words.slice(1);
}

function kebab(glyph: string): string {
  return glyph.replace(/([a-z0-9])([A-Z])/g, "$1-$2").toLowerCase();
}

let catalog: readonly OfficeIconEntry[] | null = null;

/** Every icon once, in catalog order, with all the categories it is in. */
export function officeIconCatalog(): readonly OfficeIconEntry[] {
  if (catalog) return catalog;
  const byGlyph = new Map<string, { glyph: string; categories: OfficeIconCategory[]; zh: string[]; en: string[] }>();
  for (const line of RAW.trim().split("\n")) {
    const [category, glyph, zh, en] = line.split("|");
    const entry = byGlyph.get(glyph) ?? { glyph, categories: [], zh: [], en: [] };
    if (!entry.categories.includes(category as OfficeIconCategory)) {
      entry.categories.push(category as OfficeIconCategory);
    }
    entry.zh.push(...zh.split("、"));
    entry.en.push(en);
    byGlyph.set(glyph, entry);
  }
  catalog = Object.freeze(
    [...byGlyph.values()].map((entry) =>
      Object.freeze({
        id: LEGACY_IDS[entry.glyph] ?? kebab(entry.glyph),
        glyph: entry.glyph,
        categories: Object.freeze(entry.categories),
        zh: entry.zh[0],
        en: humanize(entry.glyph),
        keywords: [...new Set([...entry.zh, ...entry.en, humanize(entry.glyph)])].join(" ").toLowerCase(),
      }),
    ),
  );
  return catalog;
}

/** The colour Insert > Icons draws an icon in (the window's and the AI's). */
export const OFFICE_ICON_DEFAULT_COLOR = "#1e293b";

/**
 * The icon a command names: its id, else its English or Chinese name, else
 * the first icon the gallery's search finds for the words. Null when the
 * catalog has none.
 */
export function resolveOfficeIcon(query: string): OfficeIconEntry | null {
  const wanted = query.trim().toLowerCase();
  if (!wanted) return null;
  const all = officeIconCatalog();
  return (
    all.find((entry) => entry.id === wanted) ??
    all.find((entry) => entry.en.toLowerCase() === wanted || entry.zh === query.trim()) ??
    searchOfficeIcons(wanted)[0] ??
    null
  );
}

/** The key a prepared icon drawing is held under: the icon and its colour. */
export function officeIconDrawingKey(id: string, color: string): string {
  return `${id}|${color.toLowerCase()}`;
}

/**
 * The icons of `category` (or every icon for `"all"`) that match every word
 * of `query`, Chinese or English.
 */
export function searchOfficeIcons(
  query: string,
  category: OfficeIconCategory | "all" = "all",
  /** The icon's name in the reader's language, searched beside the keywords. */
  localName?: (entry: OfficeIconEntry) => string,
): readonly OfficeIconEntry[] {
  const words = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
  return officeIconCatalog().filter((entry) => {
    if (category !== "all" && !entry.categories.includes(category)) return false;
    const local = localName?.(entry).toLowerCase() ?? "";
    return words.every((word) => entry.keywords.includes(word) || local.includes(word));
  });
}
