��, #�  m    LineDlg"& ' � "Communication Demo - Line Settings�  �    (   B $Form4%5�  6�  7  8(  �+   #Cancel_Command� Cancel>H7Y �#   "
OK_Command� OK�H7Y �)   !Reset_Command� Reset� H7Y �&   Poll_Buffer_Frame�x ��� �.    Poll_Interval_Text�+WText6 �2   Poll_Interval_Scroll	�2
+�	 
 2 3 �7   Poll_Interval_Label� Poll Intervalx +7 �D   Poll_Buffer_Label� Poll/Read Receive Buffer  Z � �(   Signal_Detect_Frame�x '� �1   Signal_Detect_DS_Text��WText5 �2   Signal_Detect_DS_Scroll	�A
�f	   �1   Signal_Detect_CS_Text��WText4 �2   Signal_Detect_CS_Scroll	�P
�W	   �1   Signal_Detect_CD_Text��WText3 �2   Signal_Detect_CD_Scroll	�_
�9	   �;   Signal_Detect_DS_Label� Data Set Readyx �� �:   Signal_Detect_CS_Label� Clear to Sendx �� �;   Signal_Detect_CD_Label� Carrier Detectx �7
 �C   Signal_Detect_Label� Signal Detection Time  Z �	 �0   	Comm_Buffer_Size_Frame�    x � �7  �.   Comm_Transmit_Text��WText2 �/   Comm_Transmit_Scroll	�}
�	   �-   Comm_Receive_Text��WText1 �.   Comm_Receive_Scroll	��
�	   �2   Comm_Transmit_Label� Transmitx �� �0   Comm_Receive_Label� Receivex �� �J   
Comm_Buffer_Size_Label� Communication Buffer Size  Z � ��  �  __	 
�Q�Q	   � � s�    � � -&� y �  Q"��   "  ��Z ���+�                  	  d �WorkRBn �WorkTBh�WorkDCB% CommStateDCB�WorkIntervalJ�WidthOfTextH@	Form_Load�  Remove_Items_From_SysMenu� �LineDlg�@
Initializec CenterDialogt@SizeControls;�Comm_Receive_Scroll� valueS�CommRBBuffer��Comm_Transmit_Scroll��CommTBBuffer%�CommDeviceNum��
CommHandle=�Comm_Buffer_Size_Frame enabled* FALSE��Comm_Buffer_Size_Label��Comm_Receive_Text�Comm_Receive_Label  �Comm_Transmit_Text��Comm_Transmit_LabelX�Signal_Detect_CD_Scroll��	CommState� 
RlsTimeOut��Signal_Detect_CS_Scroll� 
CtsTimeOut1�Signal_Detect_DS_Scroll% 
DsrTimeOutu�Poll_Interval_Scroll��CommReadInterval  @Reset_Command_Click�@Cancel_Command_Click  @OK_Command_Click  �NoChange> TRUE� UpdateCaption��ResultX�PostRBBuffer  �PostTBBuffer��	PostState��PostReadInterval��Receive Receive_Timer interval��ApiErr� SetCommState   DisplayQBOpen��CommPortName>@Signal_Detect_CD_Scroll_ChangeE@ProcessScrollChange`�Signal_Detect_CD_Texty@Signal_Detect_CS_Scroll_Change��Signal_Detect_CS_Texth@Signal_Detect_DS_Scroll_Change  �Signal_Detect_DS_Text�@Comm_Receive_Scroll_Change�@Comm_Transmit_Scroll_ChangeS@Poll_Interval_Scroll_Change��Poll_Interval_Text�@Poll_Interval_Text_KeyPress��KeyAscii�@ProcessTextKeyPress�@Comm_Receive_Text_KeyPress6@Comm_Receive_Text_Change  @ProcessTextChangep@Comm_Receive_Text_LostFocus  @ProcessTextLostFocus��A_Scroll�A_Text�WorkVal   Text   Max   Min  �A   SelStart@Comm_Transmit_Text_Change�@Poll_Interval_Text_Change�@Signal_Detect_CD_Text_Change�@Signal_Detect_CS_Text_Change�@Signal_Detect_DS_Text_Change��	Next_Text   	Sellength��Workw@Comm_Transmit_Text_KeyPress3@Signal_Detect_CD_Text_KeyPressU@Signal_Detect_CS_Text_KeyPress�@Signal_Detect_DS_Text_KeyPressK@Comm_Transmit_Text_LostFocus  @Poll_Interval_Text_LostFocus�@Signal_Detect_CD_Text_LostFocus�@Signal_Detect_CS_Text_LostFocus @Signal_Detect_DS_Text_LostFocus  @AdjustControlb�Signal_Detect_CD_Label  �Signal_Detect_Frame|�Signal_Detect_CS_Label��Signal_Detect_DS_Label  �Poll_Interval_Label  �Poll_Buffer_Frame  �A_Label  �A_Frame   Width   Left   height   Top   Control^        �      Z   � &     d   � 6  y   n   � F     �     � X     �   	  ��������    AdjustControl0�      X &    � ��  � ��  � ��  � ��      �  � �  � � � � � �� � �  � �  � � � �  � ���  � �  � � � �  � � � ��  � ��� �  � �    9 	  ��������
     Cancel_Command_Click0$      X  �       �      9 	  ��������     Comm_Receive_Scroll_Change0,      X  �        �$      9 	  ��������     Comm_Receive_Text_Change0Z      X  �        �$  �     " Z    n � Z  d  � $  �    9 	  ��������	     Comm_Receive_Text_KeyPress0<      X  h  E �        � � E$  Q    9 	  ��������     Comm_Receive_Text_LostFocus0,      X  �        �$  �    9 	  ��������     Comm_Transmit_Scroll_Change0,      X  �       ; �$      9 	  ��������     Comm_Transmit_Text_Change0Z      X  6       ; �$  �    ; " d     n � Z  d  � $  �    9 	  ��������	     Comm_Transmit_Text_KeyPress0<      X  �  E �       ; � % E$  Q    9 	  ��������     Comm_Transmit_Text_LostFocus0,      X  w       ; �$  �    9 	  ��������    	 Form_Load0@      X  �        � $  �    $   �     � $  �     9 	  ��������	    
 Initialize0@     X  �      w �    999991     �    $   �     +  "  + Z     S ; "  S d     c� �  t�� � � I �   + � �  + � �    + � �  + � �    + � �  +  � 8      = J " "  = s X "  = � � "  = n     � � "  � �     9 	  ��������!     OK_Command_Click0�     X        c� �  t�� � � I �    ,      Z  +� E R +    d  S� E j +    �  �� E � +    n  J = J� E � +    n  s = s� E � +    n  � = �� E � +      ,� I  0 �   0=   8     � V P  DIALOG: Change Active Settings (Yes), Post-Pone (No), Return to Dialog (Cancel)  � $  -    �  � Port Already Active!  �  � � �  � Activate settings Now?  � ��  � �  � � & �! Terminal Sampler II - Port Active '>   >V  0�&  @� " *  Changing Port Settings LIVE! � $  - @ @ Z  + @ d  S @ n  = @ �  �   @ Z  H @ d  X @ n  h @ �  u @ @ � � � �   @ n   �� @ @ �   0�&  @� , '  Settings Post-Poned until next CONNECT � $  - @ @ Z  H @ d  X @ n  h @ �  u   @ �   0%  @ n � Z  d  � $  �  :    2      Z  +   d  S   n  =   �  �     Z  H   d  X   n  h   �  u     �     8     9 	  ��������G     Poll_Interval_Scroll_Change0,      X  �       � $      9 	  ��������     Poll_Interval_Text_Change0Z      X  S       � $  �    � " �     n � Z  d  � $  �    9 	  ��������	     Poll_Interval_Text_KeyPress0<      X  &  E �       �  � E$  Q    9 	  ��������     Poll_Interval_Text_LostFocus0,      X  �       � $  �    9 	  ��������     ProcessScrollChange0h      X    � ��  � ��      � "� � �   � � � � I b   � �  8     9 	  ��������	     ProcessTextChange0�      X  �  � ��  � ��        � �     � � I ~    � � �  �    �   � � 6 �   = �  8     � � !%    � "   %� � I �    � � ! � * 8     9 	  ��������     ProcessTextKeyPress02     X & Q  � ��  � ��  � ��  E �       EV B   � &   �  &  0�  E 0 �e    x   �  �  0 � �  �  9 � '  0 � �� � I X @ E�  �  0 � � I   P � *� � I �  `�  E P � � � � 6  `�  E P8  @ � � ! � � �  �� � 6 T P�  E @8  08   %  0�  E :     9 	  ��������     ProcessTextLostFocus0�      X  �  � ��  � ��     � � !� � I r    � � �  �    � � ! � *  � �  � � 6 �    � � �  �  8    9 	  ��������
     Reset_Command_Click0$      X  �      $   �     9 	  ��������     Signal_Detect_CD_Scroll_Change0,      X  �       " %$      9 	  ��������     Signal_Detect_CD_Text_Change0^      X  p       " %$  �    " " n  J    n � Z  d  � $  �    9 	  ��������	     Signal_Detect_CD_Text_KeyPress0<      X    E �       " % ` E$  Q    9 	  ��������     Signal_Detect_CD_Text_LostFocus0,      X  �       " %$  �    9 	  ��������     Signal_Detect_CS_Scroll_Change0,      X  >       X `$      9 	  ��������     Signal_Detect_CS_Text_Change0^      X  �       X `$  �    X " n  s    n � Z  d  � $  �    9 	  ��������	     Signal_Detect_CS_Text_KeyPress0<      X  3  E �       X ` � E$  Q    9 	  ��������     Signal_Detect_CS_Text_LostFocus0,      X  �       X `$  �    9 	  ��������     Signal_Detect_DS_Scroll_Change0,      X  y       � �$      9 	  ��������     Signal_Detect_DS_Text_Change0^      X  �       � �$  �    � " n  �    n � Z  d  � $  �    9 	  ��������	     Signal_Detect_DS_Text_KeyPress0<      X  U  E �       � �  E$  Q    9 	  ��������     Signal_Detect_DS_Text_LostFocus0,      X  �       � �$  �    9 	  ��������     SizeControls0�      X  �       � � �$     ; �  �$     " % 1 K$     X ` b K$     � � | K$     �  � �$     9 	  ��������   �