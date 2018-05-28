using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using static DBApp.DatabasePreferences;

namespace DBApp
{
    public partial class UpdateWindow : Form
    {
        private Form1 formActivity;
        public UpdateWindow(Form1 forma)
        {
            InitializeComponent();
            this.formActivity = forma;
        }

        public UpdateWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (label3.Visible)
                label3.Visible = false;

            if (!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text) &&
                !string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text) &&
                !string.IsNullOrEmpty(comboBox1.Text) && !string.IsNullOrWhiteSpace(comboBox1.Text) &&
                !string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrWhiteSpace(textBox3.Text) &&
                !string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrWhiteSpace(textBox4.Text) &&
                !string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrWhiteSpace(textBox5.Text) &&
                !string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrWhiteSpace(textBox6.Text) &&
                !string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrWhiteSpace(textBox8.Text) &&
                !string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrWhiteSpace(textBox9.Text) )
            {
                SqlCommand command = new SqlCommand("UPDATE [Table] SET [Name]=@Name, [MilitaryRank]=@MilitaryRank, [Command]=@Command, " +
                    "[YearOfBirth]=@YearOfBirth, [VUS]=@VUS, [MilitaryPost]=@MilitaryPost, [TelephoneNumber]=@TelephoneNumber, [ResidentialAddress]=@ResidentialAddress, " +
                    "[PlaceOfWork]=@PlaceOfWork, [Position]=@Position, [HomePlot]=@HomePlot, [OfficialPlot]=@OfficialPlot, [HomeRoute]=@HomeRoute, [OfficialRoute]=@OfficialRoute, " +
                    "[Liter]=@Liter, [ArriveBy]=@ArriveBy,  [byVUS]=@byVUS, [JobQualification]=@JobQualification, [JQCode]=@JQCode, [MRByRegistration]=@MRByRegistration, [MRCode]=@MRCode, " +
                    "[TypeOfWeapon]=@TypeOfWeapon, [ToWCode]=@ToWCode, [RU]=@RU, [CategoryOfWorkability]=@CategoryOfWorkability, [ReserveCategory]=@ReserveCategory, " +
                    "[SpecialTalents]=@SpecialTalents, [OVKKKpoVUS]=@OVKKKpoVUS, [OVKKJobQualification]=@OVKKJobQualification, [OVKKKJQCode]=@OVKKKJQCode, [OVKKKMR]=@OVKKKMR, " +
                    "[OVKKKMRCode]=@OVKKKMRCode, [OVKKKToW]=@OVKKKToW, [OVKKKToWCode]=@OVKKKToWCode, [Education]=@Education, [MilitaryEducation]=@MilitaryEducation, " +
                    "[PersonalNumber]=@PersonalNumber, [CivilSpecialty]=@CivilSpecialty, [MilitaryUnit1]=@MilitaryUnit1, [MilitaryUnit2]=@MilitaryUnit2, [VUSCode1]=@VUSCode1, " +
                    "[VUSCode2]=@VUSCode2, [MPCode1]=@MPCode1, [MPCode2]=@MPCode2, [MRCode1]=@MRCode1, [MRCode2]=@MRCode2, [MP1]=@MP1, [MP2]=@MP2, [ToWCode1]=@ToWCode1, " +
                    "[ToWCode2]=@ToWCode2, [ToW1]=@ToW1, [ToW2]=@ToW2, [Start1]=@Start1, [Start2]=@Start2, [End1]=@End1, [End2]=@End2, [AlternateMU1]=@AlternateMU1, [AlternateMU2]=@AlternateMU2, " +
                    "[AlternatePost1]=@AlternatePost1, [AlternatePost2]=@AlternatePost2, [AlternateStart1]=@AlternateStart1, [AlternateStart2]=@AlternateStart2, [AlternateEnd1]=@AlternateEnd1, " +
                    "[AlternateEnd2]=@AlternateEnd2, [MCYear]=@MCYear, [MCAmountOfDays]=@MCAmountOfDays, [MCMU]=@MCMU, [MCVUSCode]=@MCVUSCode, [MCMRCode]=@MCMRCode, [MCMP]=@MCMP, " +
                    "[MCToWCode]=@MCToWCode, [MCToW]=@MCToW, [MaritalStatus]=@MaritalStatus, [DocumentWasDelivered]=@DocumentWasDelivered, [InsteadWho]=@InsteadWho, " +
                    "[MedicalExamination]=@MedicalExamination, [Bacteriocarrier]=@Bacteriocarrier, [Hostilites]=@Hostilites, [AccessNumber]=@AccessNumber, [SpecialNotes]=@SpecialNotes, " +
                    "[Height]=@Height, [Headdress]=@Headdress, [GasMask]=@GasMask, [Outfit]=@Outfit, [Shoes]=@Shoes, [Day]=@Day, [Month]=@Month, [Year]=@Year WHERE [Id]=@Id", getDb());

                command.Parameters.AddWithValue("Id", textBox2.Text);
                command.Parameters.AddWithValue("Name", textBox1.Text);
                command.Parameters.AddWithValue("MilitaryRank", comboBox1.Text);
                command.Parameters.AddWithValue("Command", textBox3.Text);
                command.Parameters.AddWithValue("YearOfBirth", textBox4.Text);
                command.Parameters.AddWithValue("VUS", textBox9.Text);
                command.Parameters.AddWithValue("MilitaryPost", textBox5.Text);
                command.Parameters.AddWithValue("TelephoneNumber", textBox6.Text);
                command.Parameters.AddWithValue("ResidentialAddress", textBox8.Text);
                command.Parameters.AddWithValue("PlaceOfWork", textBox7.Text);
                command.Parameters.AddWithValue("Position", textBox60.Text);

                command.Parameters.AddWithValue("HomePlot", textBox13.Text);
                command.Parameters.AddWithValue("OfficialPlot", textBox10.Text);
                command.Parameters.AddWithValue("HomeRoute", textBox11.Text);
                command.Parameters.AddWithValue("OfficialRoute", textBox12.Text);
                command.Parameters.AddWithValue("Liter", textBox62.Text);//
                command.Parameters.AddWithValue("ArriveBy", textBox61.Text);//
                //Состоит на воинском учёте
                command.Parameters.AddWithValue("byVUS", textBox15.Text);
                command.Parameters.AddWithValue("JobQualification", textBox21.Text);
                command.Parameters.AddWithValue("JQCode", textBox22.Text);
                command.Parameters.AddWithValue("MRByRegistration", textBox17.Text);
                command.Parameters.AddWithValue("MRCode", comboBox2.Text);
                command.Parameters.AddWithValue("TypeOfWeapon", textBox18.Text);
                command.Parameters.AddWithValue("ToWCode", textBox16.Text);
                command.Parameters.AddWithValue("RU", comboBox3.Text);
                command.Parameters.AddWithValue("CategoryOfWorkability", comboBox4.Text);
                command.Parameters.AddWithValue("ReserveCategory", comboBox5.Text);
                command.Parameters.AddWithValue("SpecialTalents", textBox23.Text);
                //Предназначен в ОВККК:
                command.Parameters.AddWithValue("OVKKKpoVUS", textBox30.Text);
                command.Parameters.AddWithValue("OVKKJobQualification", textBox27.Text);
                command.Parameters.AddWithValue("OVKKKJQCode", textBox26.Text);
                command.Parameters.AddWithValue("OVKKKMR", textBox32.Text);
                command.Parameters.AddWithValue("OVKKKMRCode", comboBox6.Text);
                command.Parameters.AddWithValue("OVKKKToW", textBox33.Text);
                command.Parameters.AddWithValue("OVKKKToWCode", textBox31.Text);
                //Образование, Военное образование (для офицеров), Личный номер, Осн. гражданская специальность:
                command.Parameters.AddWithValue("Education", textBox29.Text);
                command.Parameters.AddWithValue("MilitaryEducation", textBox28.Text);
                command.Parameters.AddWithValue("PersonalNumber", textBox25.Text);
                command.Parameters.AddWithValue("CivilSpecialty", textBox24.Text);
                //Срочная военная служба:
                command.Parameters.AddWithValue("MilitaryUnit1", textBox37.Text);
                command.Parameters.AddWithValue("MilitaryUnit2", textBox70.Text);
                command.Parameters.AddWithValue("VUSCode1", textBox14.Text);
                command.Parameters.AddWithValue("VUSCode2", textBox67.Text);
                command.Parameters.AddWithValue("MPCode1", textBox19.Text);
                command.Parameters.AddWithValue("MPCode2", textBox68.Text);
                command.Parameters.AddWithValue("MRCode1", comboBox7.Text);
                command.Parameters.AddWithValue("MRCode2", comboBox8.Text);
                command.Parameters.AddWithValue("MP1", textBox36.Text);
                command.Parameters.AddWithValue("MP2", textBox69.Text);
                command.Parameters.AddWithValue("ToWCode1", textBox63.Text);
                command.Parameters.AddWithValue("ToWCode2", textBox66.Text);
                command.Parameters.AddWithValue("ToW1", textBox64.Text);
                command.Parameters.AddWithValue("ToW2", textBox65.Text);
                command.Parameters.AddWithValue("Start1", textBox41.Text);
                command.Parameters.AddWithValue("Start2", textBox72.Text);
                command.Parameters.AddWithValue("End1", textBox38.Text);
                command.Parameters.AddWithValue("End2", textBox71.Text);
                //Альтернативная военная служба:
                command.Parameters.AddWithValue("AlternateMU1", textBox40.Text);
                command.Parameters.AddWithValue("AlternateMU2", textBox34.Text);
                command.Parameters.AddWithValue("AlternatePost1", textBox39.Text);
                command.Parameters.AddWithValue("AlternatePost2", textBox20.Text);
                command.Parameters.AddWithValue("AlternateStart1", textBox43.Text);
                command.Parameters.AddWithValue("AlternateStart2", textBox51.Text);
                command.Parameters.AddWithValue("AlternateEnd1", textBox42.Text);
                command.Parameters.AddWithValue("AlternateEnd2", textBox35.Text);
                //Военные сборы:
                command.Parameters.AddWithValue("MCYear", textBox45.Text);
                command.Parameters.AddWithValue("MCAmountOfDays", textBox44.Text);
                command.Parameters.AddWithValue("MCMU", textBox52.Text);
                command.Parameters.AddWithValue("MCVUSCode", textBox46.Text);
                command.Parameters.AddWithValue("MCMRCode", comboBox9.Text);
                command.Parameters.AddWithValue("MCMP", textBox48.Text);
                command.Parameters.AddWithValue("MCToWCode", textBox49.Text);
                command.Parameters.AddWithValue("MCToW", textBox47.Text);
                //Страница 2
                command.Parameters.AddWithValue("MaritalStatus", textBox56.Text);
                command.Parameters.AddWithValue("DocumentWasDelivered", textBox55.Text);
                command.Parameters.AddWithValue("InsteadWho", textBox73.Text);
                command.Parameters.AddWithValue("MedicalExamination", textBox54.Text);
                command.Parameters.AddWithValue("Bacteriocarrier", textBox53.Text);
                command.Parameters.AddWithValue("Hostilites", textBox57.Text);
                command.Parameters.AddWithValue("AccessNumber", textBox58.Text);
                command.Parameters.AddWithValue("SpecialNotes", textBox59.Text);
                //Антропометрические данные
                command.Parameters.AddWithValue("Height", textBox76.Text);//Height, Headdress, GasMask, Outfit, Shoes
                command.Parameters.AddWithValue("Headdress", textBox77.Text);
                command.Parameters.AddWithValue("GasMask", textBox78.Text);
                command.Parameters.AddWithValue("Outfit", textBox79.Text);
                command.Parameters.AddWithValue("Shoes", textBox80.Text);
                //МобПрд выдано
                command.Parameters.AddWithValue("Day", textBox81.Text);//Day, Month, Year
                command.Parameters.AddWithValue("Month", textBox82.Text);
                command.Parameters.AddWithValue("Year", textBox84.Text);

                command.ExecuteNonQuery();
                formActivity.RefreshData();
                Close();
            }
            else
            {
                label3.Visible = true;
                label3.Text = "Ошибка! Необходимо заполнить все поля!";
            }
        }
    }
}
