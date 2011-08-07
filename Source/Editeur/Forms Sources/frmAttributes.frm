VERSION 5.00
Begin VB.Form frmAttributes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attribut"
   ClientHeight    =   5190
   ClientLeft      =   1275
   ClientTop       =   1950
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub optArena_Click()
    frmArena.Show vbModal
    frmArena.scrlNum1.max = MAX_MAPS
End Sub

Private Sub OptBank_Click()
variable = InputBox("Message d'accueil:", "Banque")
bankmsg = variable
End Sub

Private Sub optCBlock_Click()
    frmBClass.scrlNum1.max = Max_Classes
    frmBClass.scrlNum2.max = Max_Classes
    frmBClass.scrlNum3.max = Max_Classes
    frmBClass.Show vbModal
End Sub

Private Sub optClassChange_Click()
    frmClassChange.scrlClass.max = Max_Classes
    frmClassChange.scrlReqClass.max = Max_Classes
    frmClassChange.Show vbModal
End Sub

Private Sub optNPC_Click()
    frmNPCSpawn.Show vbModal
    frmNPCSpawn.scrlNum.max = MAX_NPCS
End Sub

Private Sub optWarp_Click()
    frmMapWarp.Show vbModal
End Sub

Private Sub optItem_Click()
    frmMapItem.scrlItem.value = 1
    frmMapItem.Show vbModal
End Sub

Private Sub optKey_Click()
    frmMapKey.Show vbModal
End Sub

Private Sub optKeyOpen_Click()
    frmKeyOpen.Show vbModal
End Sub

Private Sub optNotice_Click()
    frmNotice.Show vbModal
End Sub

Private Sub optScripted_Click()
    frmScript.Show vbModal
End Sub

Private Sub optShop_Click()
    frmShop.scrlNum.max = MAX_SHOPS
    frmShop.Show vbModal
End Sub

Private Sub optSign_Click()
    frmSign.Show vbModal
End Sub

Private Sub optSound_Click()
    frmSound.Show vbModal
End Sub

Private Sub optSprite_Click()
    frmSpriteChange.scrlItem.max = MAX_ITEMS
    frmSpriteChange.Show vbModal
End Sub

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub

