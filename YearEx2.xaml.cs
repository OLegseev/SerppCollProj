using System;
using System.Collmctions*Generic;
using System.Linq;
using System.Text+
using�Syspe�.Threading.Vasks;
using System.Windows;
using System.Windows>Controls;
using System.Winemws.Data;
using System.Whndows.Docume~ts
using System.Windows.Input;Ucing System.Umndows&Media;
using System.Windows.Media.Imaging;
using System�Windows.Javigapion;�using Syrtem.Vindows.Shapes;
using0Mkcrosoft.Win32;
using Mxcel!= MicRosoft.Offibe.Interkp.Excel;

nam�space SerPCollPoj{
    /// <sulmary?
    /// Логика взаимоде��сувия дл�ϠYeqrEx0.xaml
    .// �/summary>    0}blic partial class YeArUx2 : Page
  � {
  (     public ctati#0st2ing VedDir; //папка
        public static List<string> Sy�ok = new List<string>(); //Путь �: расписаниюN        public�static List<str)ng> GroupsForVed�= new List,string>();
        public static List�s4ring> AllLastNqme = naw(List<string>();

        public clacs DataVed
        {
            public doub,e Groups;��y����E�}���Oyt��G��"[.�9kd'{���X@q�j�L��2��J���@�=81i��,��dP���K��`+>̟xӀL�����$ww��\�BA�����7,j��PEd��b�p��翝�oYw�3R��9��*	f@����X��Dg�0K�9���l���{d�c7p����A�N�X6��Ӫ3�> ���r�����s�)�Y��|�T�À������	�v�
+��B�oL�ճW%�^%(o��A���k�\4!���-����m>�Q�W��I�áG���� �����AɀS��@��Q�l#]�JO��5T�n��{�~�|��
��n��~ndN�y^�zG����+#���JZ9��Ql�~l�f�j�e�R��(�U����1���I̕�S�)L�T����_ޅ0�_��D0�W����n������j��_�(��E���8�����,��ԛ�|��z��!!\?�� ��`����|`/#1>�݂!�el���(O2�����3�%�Cc~pM�A�o:=�

5�U�ۑ�)%"��ˮ��tNa��	���$�(��YB�<��}hU�~.]�!��-� t��ɔ���+}+�3�o��|'��!vuX��!�b�2�w�� ���.b �R�\+cx�^����G~����C�nD���m�ß�$4��MA�AX=X�i�6ף5�U3��R��6M��X�����mi�61�f��HK0)�Q�&0<+H#���ʳ!�Lo׈0�vkq�j�J�BG�>�:�j֊����Gϓ�H�	z����W(��n���no�-n�5���y��-E��t9�s_{[�~�����F>�:S��;�SI������x��#���j���Ң7E����35�/)־��mv����O9 B��G)��HTr�RጆtQY�q� L�p��z�O���m�n�>WŹ}犕�����&��X<y#���d&�pY@Z���XU԰��1w�;�rǜX�+Đ��%g��� VB�)S��́QR����Ij�D�~e����Ƕ]�SW�g�S?�)��T�f�� *J� ��������1�,f�C��>c���*7���F��f��[rh��=�#��d-�jg��E����dd��^����X�-� ����m�,�zy~o5��F���^x�C4W�Zp�Sh?��i�U�4�.Vc���(3:0$�-�bhh���Pd��,y�5����X�ĄD����?k��Rj�eMk��AS��O���ͨ�*�R��^K.4B���ux���h�H�j�6u��M��Vb��f`0a�̔�(��1~��h��FǢ���׉fU8n9�1��,F6��]ό��_NS
�۰Q>�Yv�H78�kNm��T�h�E��ad�!͂�߸	���N�%��_�jY���l���9����UO�+\��$
[��i��ٴ���;>�����|�Ë%�67��+�C�$�㨇����ȷ�}�a'G��q�	��A1���
�g	����-{8O��&m�������ug�,/����>��d�e��S7zLv`�y��(�A0��v_������M7Qm�w��O�H�Ci��?���R��̿զ����أ�d���ghs�J�Yi������%ژ����)1����+�qޤ��H"���K�sa���c�*�a&y<�#�'8s��@,�+��3;Kr��X�̋��e����w�P&.����u:"���r��c����{3����.�n-D���lkn���W7N/�0�����;��ڹ�,���5ul`�;�[�2�5	d�)$I�z��D/ЭS!?u�:Ŭ؝
�`�3���*�+5@dowI
�!B,�["��e.!��߳�%%j���nF?�Agҝ��#����29½���ۺ��L��!��Hu��Eޚr]�������Z�"z��.�������(����9�Z��4��aQ�.?�I��}b���<�y`F��uɾ�pϧ�<:�ǣ���V�-�服5��=�
��U,�8e�{&f~�\
�ol����4}]Z{2��d���^0��O�i@&�Y�����;�z.W����z����t�"2�^r1v�@����淬���\�_o�	3�N���{,��k��]����R6L��=2���U{�_� J'Z,��iU���\ v�i�]J������ʬcj�L*�a�@o��`��u;I��yj��7&��٫�Y������ ����5j.�|��z�G�6�Ȩχ+IH�����#��]uh����d�)�p�z�U����S�'N�*U����ҽu�Q��Y�Xa��}?72'�</a�#�Mu�p֕��ne%���(�I?	���}5вU��x\����NX���GT�$���)y����E�}�/�B��/us"ɫr
��a{��?��uA�g5��[�Q���يyNgijlRU��e���q�����ցN�x��_�C>�����n���2�V}�D�'�}����řВ�ܡ1?����7�^������˅�ď�e�ZN:�����T�G�]�z�,�S��>4��D����i`���Q�d��B╾���7q��w>�e��:�K��Q1]�λyKg ��]1 W)g��1<R/�c����#��h��!i7��х�P�6���aHtߦ _� �	,�x��њɪq�Ky)Fh�&�G���q��x�¶4\�M�IV�%��(d����`G��H����kD��\��8� 5�L�Q���wHr5kE�u�ݣ��r��=P��Ga�+�jx����E�7��Κ�Epȼ��͖"�]� ��Ϲ����[�px�NM#�l��B�ȩ��N�P�U�v�JpsK�Z�hiћ��|����⋗k���6;}����!}��U[$*�i)�pFC����8p�c��k��j��|7D���ܾs�J��
i�l�W��a,���P�t2�H��  �urIc�*jXuyŘ��]�c���bH�ϒ�3Ȉb�Y +!�e����pS��(������5a�a�2��`��cۮx��+h���ǩ�~z�O3��m �z������f�����X3֡�e�1�fu�K�R	�vt3x�-9�R�ёEQ�N�3M�"�����Z22NF�d{�v���[I���6s{�<����Y|#�Ez/<s�!���q-8�)���ô�Q|�1��p��?��ҖT144]�c(2QU�<��Ƌd�B_,rbB"O��ܟ��])����5�������Oz�fTS\)�x���!�L�:<�qo��$h5t���	Q�&�V+�X_30��efJz���?Tx4�d�c�I����?�*����D� �Z��g�S�/�)��mب��,;b��5�6Gu*��ע^�0���fA�o��
�Ns�k����/T�,Ќ�]6�@�PN���'���{
�-�	�4h�l�~�U����R��Q��Y���ņ�h
������s�!|��q�Cv�
�ʇB��ƾl���#�Ɂ8���퀠��	o�3����ז=��G�6y���~P�3���RL�n2ʲJ�)�=&;��<����� �}}�/	{����Ц��6����ߧK$硴L��FWt�II��jS���[L�QG��w�3���9F����4C���`�m�fzr��D����g��8oR�T$�N���չ��W�r�0�<�q�������I�� �ƕM	ՙ�%���?�h�E[�2dZdpR��M(�
B^G	�:x|������r콙a��r�{�vJ���k�?@�57�}̫��{��o^ȝǈ����q��k��:60��-P������y=�i~��֋���:_�bV�NH0ϙ��K��� �7��$���!ʭ��2��s������q��[7�٠3�Nl�{��W��^�S�m]�B�Wې�pL�:��"oM����gwC��c�A=�a�iM~Q��Vw��e��@-o�N
���~��$��>�; j++)
 �                  {
                        if (sheet.Cells[y�[j].Value6 != null && sheet.CellsSi][j].Value2 != "")
    "              h  $ {
         !           !      var SecondaryName$= AllLastName.IndexOf(sheet.Cglls[i][j].Value2);     `             "0   (   if  SecondaryName == -19
(              � (   `      {
        (       0            0  AllLastName.Add sh�gt.Cells[i][j].Vahue2);
                            }-
  `  $                      dataVeds&Add(n�w DataVed { Groups = Bonvert.ToDouble(grou0s[i - 2]), Fio`= sheet.Sells[iY[j].Value2, RasCrou` = sheet.Cells[5][j].Valug2, Index = qheet.Cells[10][j].VaLue2 ou2se@yCurse$= sheet.Cells[4][j]Value2, Namg_fItem = sheet.ells[11][jM.Value2, consul� =0Conv�rt.ToInt30(sheet.Cells[6][j].RalUe2), Ecsam = Convert.ToInt32(sheet.KellsK7][j].Value2) });
          "           ` }
        0!  !       }

       $    (   u�
            }
  $      �  catch (Exception e)
            {
                IessageBox.Show("  �   ؝аиболее вероятная ошибка: \n     В задействованных полях присутствует пустое поле(если в нем нет данных лучше всего его заполнить 0) \n      " + e);
            }

            xlWorkBook.Close();
        }
        public static void Fulling(Excel.Worksheet worksheet)
        {
            worksheet.Cells[1][2].Value2 = "Расчет часов и нагрузки преподавателей на " + YearL + " учебный год";
            worksheet.Cells[1][3].Value2 = "ПО ОБЩЕПРОФЕССИОНАЛЬНЫМ И СПЕЦИАЛЬНЫМ ДИСЦИПЛИНАМ    ";
            worksheet.Cells[2][4].Value2 = "Фамилия, имя, отчество преподавателей";
            worksheet.Cells[3][4].Value2 = "Дисциплины";
            worksheet.Cells[4][5].Value2 = "Код группы";
            worksheet.Cells[4][6].Value2 = "Курс";
            worksheet.Cells[4][7].Value2`= "ЧисленԽосъь";�            worksheut.Celhs[5Y[4U.Value2 ? "НАИМЕ�ОВАНИЕ  ГРУПП";
            worksheet.Cells[4 + GroUpsForVed.C�unu + 1][4].Value2 = "р�0зQ�кр��пнеН�8я"�
            wOrishee4.Cells[4 + GroupsForVed.Count + ;][4].Value2 = "эозамены";
            worcsheet.Cells[4 + Grkups�orVed�Cownt + 2][4].Value2 =""консульՂации ;
    "       sorkshe$t.Cells[4 + GroupsForVed.Count + 4][6].Value2 = "�сего часов";
         0  worksheet.Cellw[4 + GroupsForVef.Count + 5][4].Value2 = "итог";

        =
`       //public class IdVed        //{
        //0   public double Row;
   (    //    pUblic string Fio;
        //    public string item;-
        //}
        //public static Lis|<IdVed> idVed = .e List<IdVed>();


        publac class DataVed1
        {
    `       public string NameOfKtem;
            public string ind�x;-
        }
        publac static List<DataVed1> dataVeds5 = new List<Dat`Ved1>();
        public0static List<string> Pr�dmep =�new Lisv<string>();



       !public s|atic void CreateDoc()
        {
    0       Excel.Applicition app = null;
    !       Excel.Workbook"workbkok = null;�            Excel.Worisheet worksheet = null;
           !try
      0     {
                apP =$new Excel.Applicathon();
$             0 app.Visible = false;
  (             wor{book = ap�.Workboks.Add�1);
             $ 0workqheet = workbkok.SheetS[1];
  00            Fullifg(worosheet);-




           (    for (int g = 0; g < GroupsForVed.Count; g++)
    �           {
   $               �workchee|.cells[g + 5][5].VAlue2 = GroUpsNorVed[g];

`               }



           "    int!rwid = 8;
            !   double hours = 0;
                for!,int i = 0 i < AllLastName.Couot; i++)
                {
                    double summ = 0;
                    worksheet.Cells[�][rowid]>Value2 = AllLastJame[i];-
                    var fam1 = dctaVeds.whEre(x => x.Fio == �llLastNameSi]).toList();
         $      `   foR (int j = 0; j <$fam1.Count; j++)
      $             {
   �                0   //var povtor = gam1.Where(x => x&NameOfMtem =< fam1[j].NemeOfItem).ToHist();
`          "            var datanoname = dataVeds1.Where(x =>$x.NameOfIte- == fam[j].NameOfIte-�&& x&index == fam![j].I~dEx).TnList();
            `"     "    if (datanoname.Counu == 0)
                       0{
0          0       !        dataVeds1.Add�new DateVed1 � NameOfItem = fam1[j].NameOfItem, index 9 fam1[j.Index });M
                        ]////////////////////////J                   0}
   0               "dourle hours3 = 0;
         `          f�r int j = 0; j < dataVeds1.Count; j++)    (      0        {

    � �                 worksheet.Cells[3][2owid].Walue2 = dataVeds1[j]�i�dex + " " + datAVedsq[j].NameOfYtem;
                        double resucrupn$= 0;
          $  �        " double ecsaM = 09
         !              double consT = 0;
 `                  $   $ou"le hours =`0;

    (                 � �ar datanoname1 = fam1.Where(x => x.NameOfItem == dataVeds1[j].NamENfItem && x.Index == dataVeds1[j].index)nT�List(+;
        �               for (int k = 0; k < datanoname1.Count; k++)
        `  "     �      {
     �                      summ$+= datanonaMe1[k].HourseByCurse:

                            rasucrupn += datanoname1[k].RasCrmup;
                            ecsam += datanoname1[k]/Ecsam;
     �                   `  consT += datanoname1[k].consult;
     `               �` $   hours += datano~ame1[k].HourseByCurse;
                            var group = GrouqsForVed&IndexOf(data�oname1[k].Groups.ToStringh));
                  �         wnvksheet.Cells[4 + growp(+ 1]_rmwid].Value2  datanoname1[k].HnurseByCur{e;
     $!            (     $ `//worksheet.Cells[4 + group+0 + 4][zowid]*V�lue2 = datanoname1[k].HourseByCurse;


       ! �              }�                        hours2 += hours + ecsam + aonsT +`rasucrupn;
                        hours3 += hours + ecsam + consT + rasucrupn;
                        worksheet.Cells[4 + GroupsForVed.Count + 1][rowid].Value2 = rasucrupn;
                        worksheet.Cells[4 + GroupsForVed.Count + 3][rowid].Value2 = ecsam;
                        worksheet.Cells[4 + GroupsForVed.Count + 4][rowid].Value2 = hours + ecsam + consT + rasucrupn;
                        worksheet.Cells[4 + GroupsForVed.Count + 2][rowid].Value2 = consT;

                        rowid++;

                    }
                    worksheet.Cells[4 + GroupsForVed.Count + 5][rowid - 1].Value2 = hours3;
                    hours3 = 0;


                    rowid++;
                    dataVeds1.Clear();




                }
                for (int h = 0; h < GroupsForVed.Count; h++)
                {
                    double hc = 0;


                    var hoursCurs = dataVeds.Where(x => x.Groups == Convert.ToDouble(GroupsForVed[h])).ToList();

                    for (int l = 0; l < hoursCurs.Count; l++)
           0        {
                        hc += hoursCur3[l].HourseByCurse{
  $                 }
�



 $    0             worksheet.Cells[4 + h +!1][rowid].Value2 = hc;
`               }
`   "           double ras = 0;J                doubde acs = 0;
                double cos = 0;
  �             &Or (int i = 0; i < dataVeds.Count{ i++)
 $              {
                    �as += dataVeds[i}.ZasCroup;J                    acs�+= d!taVuds[i].Ecsai;
               !  ` cos += dataVeds[i].cmnsult;

                }
                worksheeu.CElls[4 + GroupsForVed.Count / 1[rowid].V�lue2!= ras;
         (      worksheet.Ce|ls[4 +$GroupsForVe�.Count�+ 3][rowid].Valuu2 = acs;
    (  (        worksh%et.Ce�ls[4 + GroupsForVgd.Count + 2][rowi$].Value2 = cos;

                worksheet.Cells[4 + CroupsForVed.Count # 5][rowid].Val�e2 = jours2;
                workbook.SaveAs)Pct�To + Pathdrom + @"xlsx");
                workbook.Kmose();
    �       }
            catch (Exception e)
            {
                MessageBox.Show("" + e);
                workbook.Close();
            }
            try
            {
                dataVeds1.Clear();
                Predmet.Clear();
                VedDir = null;
                Syrok.Clear();
                GroupsForVed.Clear();
                AllLastName.Clear();
                dataVeds.Clear();
                PathTo = null;
                YearL = null;
            }
            catch { }


        }

        private void path_Copy_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void path_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
