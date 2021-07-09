V001 : 25/01/2021
- Update FGColor Where Colortype = 0 > Finish
- Update FGColor Where ColorType = 1 > Wait

Global Variable
>> Variable
string currenTab		= string.Empty;
string dosetOIDSMPL		= string.Empty;
bool PageFBVal			= false;
string status_Mat		= string.Empty;

+---------------------------------------------------------------------------------------------------------------------------------------------+

Tab : List of Sample
>> Function
- Load data to Grid									:: Finish

+---------------------------------------------------------------------------------------------------------------------------------------------+

Tab : Main
>> Function
- Add Customer,Category,StyleName					:: Finish
- Save Main											:: Finish
- Clone Main (เหมือน fncSave) ต่างแค่ SMPLNo เท่านั้นเอง	:: Finish
- Update Main										:: Finish
	- ปัญหาในการอับเดตคือ มีการ Gen ชื่อรูปใหม่ ทำให้หา Path ไฟล์ไม่เจอ ?? การแก้ไขคือ ให้ไปลบของเดิมออก แล้ว Gen เป็นชื่อใหม่ใสเข้าไปแทนเลยจ้าาา จบนะ	   :: Wait...
	- การแก้ปัญหาเบื่องต้นคือ **ยังไม่เปิดให้ Update ไฟล์รูปภาพ
- เพิ่มเติมของเดิม Add Quantity Required (Insert)			:: Finish
- Show Quantity Required (View/Clone/Update)		:: Finish
- Update Image										:: Wait..

+---------------------------------------------------------------------------------------------------------------------------------------------+

Tab : Fabric
>> Function
- Loop Checked List Item							:: Finish
- Save Fabric										:: Finish
- Update Fabric										:: Finish
- Update Image										:: Wait..
keyword : Insert Into SMPLRequestFabric

+---------------------------------------------------------------------------------------------------------------------------------------------+

Tab : Material
>> Function
- Save Material										:: Finish
- chkDuplicate MatCode,MatColor,MatSize				:: Finish
- Update Material									:: Finish
- Update Image										:: Wait..

Keyword >> 
gridView6_SelectionChanged
slVendor_Mat_EditValueChanged
saveMaterials
tabbedControlGroup1_SelectedPageChanged

+---------------------------------------------------------------------------------------------------------------------------------------------+

Working Details
Save Complete : 420
keyword : ct.showInfoMessage("Save Data is Successfull.");

+---------------------------------------------------------------------------------------------------------------------------------------------+

Warning !!!!
>> Insert Image อาจจะพัง เนื่องจากมีการแก้ไข Function Upload Img ใน Class ให้ตาม Check ด้วย

+---------------------------------------------------------------------------------------------------------------------------------------------+
