   M S B A M L   `   `   `  ����U  QLasikVSSmileInHighAstigmat, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null]3clr-namespace:LasikVSSmileInHighAstigmat.ViewModels%LasikVSSmileInHighAstigmat.ViewModels  5   	   +    %LasikVSSmileInHighAstigmat.MainWindow   5      W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   T PMahApps.Metro, Version=2.0.0.0, Culture=neutral, PublicKeyToken=51482d6f650b2b3f;Controls,http://metro.mahapps.com/winfx/xaml/controls  5   	   HviewModelLocal3clr-namespace:LasikVSSmileInHighAstigmat.ViewModels  5   	   -   5        (  "MahApps.Metro.Controls.MetroWindow   Name$  DetailsWindow��5       � Title$% Lasik Vs Smile In High Astigmat��5	   	   $	��500��61   $
��1000��6>    � 
ResizeMode$ NoResize=�6K    � DataContext 5
      9   3LasikVSSmileInHighAstigmat.ViewModels.MainViewModel  5   
   .��5      � .+�5   
   �� $
��Wrap=�6   $	��350��6(   $
��Left=�64   $��30,30,0,125q�6O   $��14(�6d   �� 5      $
��Bold"�6   Oluşturulacak Grafikler :6E   �� 6P   �� 6\   201. Vektörel analiz tabloları: TIA, SIA, DV, CI�� 5   <   2. Efficacy�� 5      3. UDVA vs CDVA�� 5      4. Safety (Change in CDVA)�� 5   (   :85. Attempted vs Achieved Spherical Equivalent Refraction�� 5   F   0.6. Accuracy of Spherical Equivalent Refraction�� 5   <   1/7. Stability of Spherical Equivalent Refraction�� 5   =   8. Amplitude of Astigmatism�� 5   )   9. TIA vs SIA�� 5      10. Angle of Error�� 5       5      �� 5   
   $
��Wrap=�6   $	��350��6(   $
��Left=�64   $��Bottom=�6O   $��	30,0,0,25q�6j   $��14(�6}   �� 5      $
��Bold"�6   Oluşturulacak Tablolar :6D   �� 6O   �� 6[   ki1. Preop Data detayları ve eğer iki grubunuz varsa ve bunların "p" değeri ile karşılaştırılması�� 5   m   lj2. Postop Data detayları ve eğer iki grubunuz varsa ve bunların "p" değeri ile karşılaştırılması�� 5    n   yw3. Vektörel Analiz Datası detayları ve eğer iki grubunuz varsa ve bunların "p" değeri ile karşılaştırılması�� 5!   y   5"      � 5$   
   $��Right=�6   $	��Top=�6+   $	��500��6C   $	��275��6O    � ColumnDefinitions 5%      �� 5&      �� 5'      .+�5*      � $��0��6   $��15q�6#   $��Stretch=�6/    � RowDefinitions 5+      �� 5,      $
��Auto �6$   �� 5-      $��* �6$   .+�50      �� -   5         ��Name$ groupsListBox��50   (   $��1��6   $��Stretch=�6?   $��Stretch=�6]   ��6y   �� *0.DefaultTemplateCreator.DataTemplate.GroupNames+ �� View 51      �� 52       �� ColumnHeaderToolTip$ Authors��6#   .P�53      �� $�� ��6-   	 �� Width$		 200��67   
 �� CellTemplate
 54   "   �� 55   &   .R�56   *   � $��0��6/   $��Center=�6<   $
��Left=�6W   $	��150��6r   .+�57   .   ��  ��Name$ GroupNameTextBox��66   $��5q�6P    �� TextAlignment$ Center=�6[   ��6r   ��  �� Path$
 Name��� 58   .    J� IsHitTestVisible$ False��68   $��
Group Name��6Q   $��Center=�6c   $��Center=�6~   $��0,0,15,0q�6�   
���������59   2   �� 5:   6    �� 
TargetType" ��6<   .�5;   :   ��  �� Property � y� �� Value$ 	Collapsed=�6W    �� Triggers 5<   :   �� 5=   >    �� Binding 6J   �� *Text+ �� ElementName$ GroupNameTextBox�� �� Value  6�   .a�5>   B   ��  � y�$ Visible=�6_   5?   ?   5A   7   5C   /   5D   +   5E   '   5G      5H      5J      �� 5L      $��Groups��6   $��5q�6*   �� 5M      $
��Ekle��6   $��Right=�6(   $	��Top=�6D   $��0,0,50,0q�6\   ��6n   �� *(&DefaultTemplateCreator.AddGroupCommand+�� 5N      $	��Sil��6   $��Right=�6'   $	��Top=�6C   $��0,0,10,0q�6[   ��6m   �� *(&DefaultTemplateCreator.DelGroupCommand+��6�   �� $ groupsListBox��$ SelectedIndex�5O      � 5Q      $��1��6   $��15q�6#   $��Stretch=�6/   $
��Left=�6K    5R      �� 5S      $
��Auto �6$   �� 5T      $��* �6$   .+�5W      �� -   5        $ periotsListBox��5W   (   $��1��6   $��Stretch=�6@   $��Stretch=�6^   $��5,0,5,5q�6z   ��6�   �� *31DefaultTemplateCreator.DataTemplate.ControlMonths+ 5X      �� 5Y      $ Authors��6#   .P�5Z      �� $�� ��6-   $		 200��67   
 5[   "   �� 5\   &   .R�5]   *   � .+�5^   .   *  $MahApps.Metro.Controls.NumericUpDown  $��3q�6b   $��Center=�6m   $
��Left=�6�   $��75��6�      HideUpDownButtons$
 True��6�      Value 6E   �� $ Month�6�   �� 5_   .   $��. Month��68   $
��Left=�6G   $��Center=�6b   $��80,0,0,0q�6}   5`   +   5a   '   5c      5d      5f      �� 5h      $��Periots��6   $��5q�6+   �� 5i      $
��Ekle��6   $��Right=�6(   $	��Top=�6D   $��0,0,50,0q�6\   ��6n   �� */-DefaultTemplateCreator.AddControlMonthCommand+�� 5j      $	��Sil��6   $��Right=�6'   $	��Top=�6C   $��0,0,10,0q�6[   ��6m   �� */-DefaultTemplateCreator.DelControlMonthCommand+��6�   �� $ periotsListBox��$ SelectedIndex�5k      5l      �� 5n   
   $��Excel Şablonu Oluştur��6   $��30,0,30,140q�61   $��Bottom=�6F   $	��400��6a   $��50��6m   $��Right=�6y   $��19(�6�   ��6�   �� *!CreateExampleDefaultDataCommand+�� 5o   
   $2��,Gelişmiş Ayarlarla Excel Şablonu Oluştur��6   $��0,0,100,100q�6D   $��Bottom=�6Y   $	��250��6t   $��35��6�   $��Right=�6�   $��10(�6�   ��6�   �� *CreateExampleDataCommand+�� 5p   
   $��Excel Şablonu Analiz Et��6   $��
30,0,30,25q�63   $��Bottom=�6G   $	��400��6b   $��50��6n   $��Right=�6z   $��17(�6�   ��6�   �� *OpenFileCommand+5r      5s      