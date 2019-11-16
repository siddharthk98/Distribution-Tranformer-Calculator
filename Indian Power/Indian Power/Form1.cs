using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Indian_Power
{
	
	public partial class Form1 : Form
	{
		String Ref_Standard;
		double Rcap,Flux_Density;
		double inC, outC, cap, inV, outV,single_phase_v,flux_constant,impedance, No_Of_Tappings;
		double Sel_CTA, ELA, Dia_sqmm, Dia_sqinch, Turn_per_volt, Final_Turn_per_volt;
		double LV_turns, HV_turns, Final_LV_turns, Final_HV_turns;
		double Conductor_Density,LV_Conductor_Area, HV_Conductor_Area, No_of_Conductors, LV_Conductor_Size, HV_Conductor_Dia;
		double LV_Conductor_Thickness, LV_Conductor_Width,LV_Insulation_Covering, HV_Insulation_Covering;
		double LV_Conductor_To_Use_Result_W, LV_Conductor_To_Use_Result_D, HV_Conductor_To_Use_Result;
		double Bare_Conductor_Size_W, Bare_Conductor_Size_D, total_LT_Turns;
		double No_Of_Layers,Turns_Per_Layer,No_Of_Conductors_W, No_Of_Conductors_D;
		double Length_Of_One_Turn, Gap_Factor, Length_Of_Winding;
		double Chip_Size_U, Chip_Size_L,Total_Length_Coil,Total_Length_Of_Winding;
		double Window_Padding_U, Window_Padding_L, Window_Height_Core;
		double Core_Inner_Dia, LV_Coil_Inner_Dia, Gap_Core_LV,LV_Paper_Gap_Factor, Outter_Coil_Dia, Total_Outter_Coil_Dia;
		double HV_Coil_Inner_Dia, Oil_Duct_Size,No_Of_Selection_HV_coils,Packing_Of_HV_Coils;
		double XL_Length_Of_HV_Window, HV_Conductor_Size, Tapping_Range_Max, Tapping_Range_Min;
		double Total_HV_Turns, HV_Total_V, Total_Turns_Single_Disc, Total_Turns_Single_Disc_RO, Turn_Per_Layer_Single_Disc, Total_Layers_Single_Disc;
		double HV_Paper_Gap_Factor, HV_Outter_Dia, Total_Layers_Single_Disc_RO, Total_Coil_Outter_Dia_RO, HV_Outter_Dia_RO;
		double Gap_of_Coils, CL_Core,Final_Total_Turns;
		double Gap_HV_Coil_Tank, Channel_H, Channel_W, Channel_Size, Foot_Plate, Foot_Plate_W, Foot_Plate_H, Height_Transformer;
		double Tank_Length, Size_Of_Tap_Changer, Tank_Upper_Gap, Tank_Lower_Gap, Tank_Width, Tank_Height;
		double Turns_Mid_C1, Turns_Mid_C2, Turns_Mid_C3, Turns_Mid_C4, Turns_Mid_C5,Tapping_Per_Step, Total_Turns_Of_Coil,Turns_Per_Tapping;
		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			Flux_Density = Double.Parse(comboBox_Flux_Density.Text.ToString());
		}

		private void TabPage1_Click(object sender, EventArgs e)
		{

		}

		private void GroupBox1_Enter(object sender, EventArgs e)
		{

		}

		private void Label3_Click(object sender, EventArgs e)
		{

		}

		private void TextBox1_TextChanged(object sender, EventArgs e)
		{
			text_Job_No.Focus();
		}

		private void Label13_Click(object sender, EventArgs e)
		{

		}

		private void TextBox7_TextChanged(object sender, EventArgs e)
		{

		}

		private void ComboBox3_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void Label21_Click(object sender, EventArgs e)
		{

		}

		private void ComboBox9_SelectedIndexChanged(object sender, EventArgs e)
		{
			No_Of_Tappings = int.Parse(combo_No_Of_Tappings.SelectedItem.ToString());
			text_No_Of_Tappings.Text = (Convert.ToDouble(No_Of_Tappings)).ToString();
		}

		private void Label42_Click(object sender, EventArgs e)
		{

		}

		private void SplitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
		{

		}

		private void TabPage8_Click(object sender, EventArgs e)
		{

		}

		private void Label89_Click(object sender, EventArgs e)
		{

		}

		private void TabPage4_Click(object sender, EventArgs e)
		{

		}

		private void CALCULATORToolStripMenuItem_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start("calc");
		}

		private void TextBox11_TextChanged(object sender, EventArgs e)
		{
			
		}

		private void OPENToolStripMenuItem_Click(object sender, EventArgs e)
		{
			openFileDialog1.ShowDialog();
		}

		private void Text_inpV_TextChanged(object sender, EventArgs e)
		{
			if (text_inpV.Text.Trim() == string.Empty)
			{
				text_inpV.Text = "";
				return;
			}
			else
			{
				inV = int.Parse(text_inpV.Text);
				inC = ((cap * 1000) / inV) / Math.Sqrt(3);
				text_inpA.Text = (Convert.ToDouble(inC)).ToString("0.00");
				
				text_Ref_Voltage.Text = (Convert.ToDouble(inV)).ToString();
			}
		}

		private void Text_outV_TextChanged(object sender, EventArgs e)
		{
			if (text_outV.Text.Trim() == string.Empty)
			{
				text_outV.Text = "";
				return;
			}
			else
			{
				outV = int.Parse(text_outV.Text);
				outC = ((cap * 1000) / outV) / Math.Sqrt(3);
				text_outA.Text = (Convert.ToDouble(outC)).ToString("0.00");
				single_phase_v = Math.Round(outV / ((Math.Sqrt(3))));
				text_Single_Phase_V.Text = (Convert.ToDouble(single_phase_v)).ToString("0.0");
			}
		}

		private void Combo_Tank_Lower_Gap_SelectedIndexChanged(object sender, EventArgs e)
		{
			Tank_Upper_Gap = int.Parse(combo_Tank_Upper_Gap.SelectedItem.ToString());
			Tank_Lower_Gap = int.Parse(combo_Tank_Lower_Gap.SelectedItem.ToString());
			Size_Of_Tap_Changer = int.Parse(combo_Size_Of_Tap_Changer.SelectedItem.ToString());
			Tank_Height = Tank_Upper_Gap + Tank_Lower_Gap + Size_Of_Tap_Changer + Height_Transformer; 
			text_Tank_Height.Text = (Convert.ToDouble(Tank_Height)).ToString("0.00");
			text_Tank_H.Text = (Convert.ToDouble(Tank_Height)).ToString("0.00");
			text_Tank_L.Text = (Convert.ToDouble(Tank_Length)).ToString("0.00");
			text_Tank_W.Text = (Convert.ToDouble(Tank_Width)).ToString("0.00");
		}

		private void Label112_Click(object sender, EventArgs e)
		{

		}

		private void TabPage5_Click(object sender, EventArgs e)
		{

		}

		private void Combo_Gap_Factor_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (combo_Tapping_Step_Value.SelectedItem == null)
			{
				Gap_Factor = 1.25;
				return;
			}
			Gap_Factor = Double.Parse(combo_Gap_Factor.SelectedItem.ToString());
			Total_Length_Of_Winding = Length_Of_Winding + ((Length_Of_Winding * Gap_Factor) / 100);
			text_Total_Length_Winding.Text = (Convert.ToDouble(Total_Length_Of_Winding)).ToString("0.00");
		}

		private void Text_Chip_Size_U_TextChanged(object sender, EventArgs e)
		{

		}

		private void PictureBox5_Click(object sender, EventArgs e)
		{
			//MessageBox.Show(tabControl1.SelectedIndex.ToString());
			tabControl1.SelectedIndex++;
		}

		private void PictureBox4_Click(object sender, EventArgs e)
		{
			//MessageBox.Show(tabControl1.SelectedIndex.ToString());
			if (tabControl1.SelectedIndex != 0)
			{
				tabControl1.SelectedIndex--;
			}
			

		}

		private void Text_capacity_KeyUp(object sender, KeyEventArgs e)
		{

			
		}

		private void Text_capacity_KeyDown_1(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Tab)
			{
				text_inpV.Focus();
			}
		}

		private void Text_Ref_Voltage_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Conductor_Thickness_TextChanged(object sender, EventArgs e)
		{

		}

		private void PictureBox9_Click(object sender, EventArgs e)
		{

		}

		private void Text_Final_HV_Turns_TextChanged(object sender, EventArgs e)
		{

		}

		private void PictureBox7_MouseClick(object sender, MouseEventArgs e)
		{
			Application.Restart();
		}

		private void GroupBox2_Enter(object sender, EventArgs e)
		{

		}

		private void Text_Flux_Constant_KeyDown(object sender, KeyEventArgs e)
		{
		}

		private void Text_Total_Coil_Outter_Dia_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Total_Coil_Outter_Dia_KeyDown(object sender, KeyEventArgs e)
		{
		}

		private void Text_LV_Turns_KeyDown(object sender, KeyEventArgs e)
		{

		}

		private void Text_HV_Turns_KeyDown(object sender, KeyEventArgs e)
		{

		}

		private void Text_Final_LV_Turns_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				Final_LV_turns = Double.Parse(text_Final_LV_Turns.Text);
				text_Final_LV_Turns.Text = (Convert.ToDouble(Final_LV_turns)).ToString("0.00");
			}
		}

		private void Text_Final_HV_Turns_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				Final_HV_turns = Double.Parse(text_Final_HV_Turns.Text);
				text_Final_HV_Turns.Text = (Convert.ToDouble(Final_HV_turns)).ToString("0.00");
			}
		}

		private void Text_Total_Coil_Outter_Dia_RO_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				Total_Coil_Outter_Dia_RO = Double.Parse(text_Total_Coil_Outter_Dia.Text);
				text_Total_Coil_Outter_Dia_RO.Text = (Convert.ToDouble(Total_Outter_Coil_Dia)).ToString("0.00");
			}
		}

		private void Text_Total_Turns_Single_Disc_RO_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				Total_Turns_Single_Disc_RO = Double.Parse(text_Total_Turns_Single_Disc_RO.Text);
				text_Total_Turns_Single_Disc_RO.Text = (Convert.ToDouble(Total_Turns_Single_Disc_RO)).ToString("0.00");
			}
		}

		private void Text_HV_Outter_Dia_RO_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_HV_Outter_Dia_RO_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				HV_Outter_Dia_RO = Double.Parse(text_HV_Outter_Dia_RO.Text);
				text_HV_Outter_Dia_RO.Text = (Convert.ToDouble(HV_Outter_Dia_RO)).ToString("0.00");
			}
		}

		private void Combo_Ref_Standard_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Tab)
			{
				text_capacity.Focus();
			}
		}

		private void Text_Job_No_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Tab)
			{
				text_capacity.Focus();
			}
		}

		private void PictureBox8_Click(object sender, EventArgs e)
		{

		}

		private void PictureBox8_MouseClick(object sender, MouseEventArgs e)
		{
			Document doc = new Document(iTextSharp.text.PageSize.A4, 49, 49, 49, 49);
			PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream("Final Result PDF.pdf", FileMode.Create));
			doc.Open();
	
			iTextSharp.text.Image JPG = iTextSharp.text.Image.GetInstance("letterhead.png");
			JPG.Alignment = Element.ALIGN_CENTER;
			JPG.ScalePercent(30f);
			doc.Add(JPG);

			iTextSharp.text.Font pfont1 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(),
			15, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

			DateTime dt = new DateTime();
			Paragraph date1 = new Paragraph("Date : " + dt.ToString("dd.MM.yyyy"));
			Paragraph details = new Paragraph("", pfont1);

			doc.Add(date1);

			Paragraph title = new Paragraph("\nDISTRIBUTION TRANFORMER JOB CARD\n\n",pfont1);
			title.Alignment = Element.ALIGN_CENTER;
			doc.Add(title);

			BaseFont bfTimes = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);

			iTextSharp.text.Font heading = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

			iTextSharp.text.Font tab = new iTextSharp.text.Font(bfTimes, 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);

			iTextSharp.text.Font title1 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED);

			PdfPTable table = new PdfPTable(4)
			{ HorizontalAlignment = Element.ALIGN_CENTER, WidthPercentage = 100, HeaderRows = 1};

			float[] widths = new float[] { 20f, 80f, 60f, 60f };
			table.SetWidths(widths);

			PdfPCell WindingAndCoil = new PdfPCell(new Phrase("WINDING & COILS",title1));
			WindingAndCoil.Colspan = 4;
			WindingAndCoil.HorizontalAlignment = 1;
			table.AddCell(WindingAndCoil);

			table.AddCell(new PdfPCell(new Phrase("SNo.", heading)));
			table.AddCell((new PdfPCell(new Phrase("PARTICULARS", heading))));
			table.AddCell((new PdfPCell(new Phrase("LV", heading))));
			table.AddCell((new PdfPCell(new Phrase("HV", heading))));

			table.AddCell(new PdfPCell(new Phrase("1.", heading)));
			table.AddCell(new PdfPCell(new Phrase("TYPE OF COIL", tab)));
			table.AddCell(new PdfPCell(new Phrase("SPIRAL", tab)));
			table.AddCell(new PdfPCell(new Phrase("CROSSOVER", tab)));

			table.AddCell(new PdfPCell(new Phrase("2.",heading)));
			table.AddCell(new PdfPCell(new Phrase("CONDUCTOR SIZE BARE",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("3.",heading)));
			table.AddCell(new PdfPCell(new Phrase("CONDUCTOR SIZE BARE",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("4.",heading)));
			table.AddCell(new PdfPCell(new Phrase("COVERING",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(" ");
			table.AddCell(new PdfPCell(new Phrase("CONDUCTOR SIZE AFTER COVERING",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("5.",heading)));
			table.AddCell(new PdfPCell(new Phrase("WINDING DIRECTION",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("6.",heading)));
			table.AddCell(new PdfPCell(new Phrase("COIL INTERNAL DIA",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("7.",heading)));
			table.AddCell(new PdfPCell(new Phrase("COIL OUTTER DIA", tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("8.",heading)));
			table.AddCell(new PdfPCell(new Phrase("CONDUCTOR HEIGHT IN COIL", tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("9.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TOTAL COIL HEIGHT WITH CHIP SIDE", tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("10.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TOTAL TURNS", tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("11.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TURNS PER LAYER", tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("12.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TOTAL LAYERS", tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("13.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TRANSPOSITION", tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("14.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TAPPINGS",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("15.",heading)));
			table.AddCell(new PdfPCell(new Phrase("WEIGHT(APPROX.)",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("16.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TOTAL WEIGTH LT + HT",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			PdfPCell Tank = new PdfPCell(new Phrase("TANK", title1));
			//cell.BackgroundColor = new iTextSharp.text.BaseColor(0, 150, 0);
			Tank.Colspan = 4;
			Tank.HorizontalAlignment = 1;
			table.AddCell(Tank);
			
			table.AddCell(new PdfPCell(new Phrase("1.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TANK SIZE WITHOUT RADIATORS", tab)));
			table.AddCell(new PdfPCell(new Phrase("LxBxH", tab)));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("2.", heading)));
			table.AddCell(new PdfPCell(new Phrase("CONSORVATOR SIZE", tab)));
			table.AddCell(new PdfPCell(new Phrase("DIAxL",tab)));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("3.",heading)));
			table.AddCell(new PdfPCell(new Phrase("OVERALL SIZE",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("4.",heading)));
			table.AddCell(new PdfPCell(new Phrase("RADIATORS",tab)));
			table.AddCell(new PdfPCell(new Phrase("FRONT",tab)));
			table.AddCell(new PdfPCell(new Phrase("BACK",tab)));

			table.AddCell(new PdfPCell(new Phrase("5.",heading)));
			table.AddCell(new PdfPCell(new Phrase("SIDE SHEET",tab)));
			table.AddCell(new PdfPCell(new Phrase("BOTTOM SHEET",tab)));
			table.AddCell(new PdfPCell(new Phrase("TOP SHEET",tab)));

			table.AddCell(new PdfPCell(new Phrase("6.",heading)));
			table.AddCell(new PdfPCell(new Phrase("LOCKING",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("7.",heading)));
			table.AddCell(new PdfPCell(new Phrase("EXPLOSION VENT",tab)));
			table.AddCell(new PdfPCell(new Phrase("DIA",tab)));
			table.AddCell(new PdfPCell(new Phrase("HEIGHT",tab)));

			table.AddCell(new PdfPCell(new Phrase("8.",heading)));
			table.AddCell(new PdfPCell(new Phrase("SPACER",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("9.",heading)));
			table.AddCell(new PdfPCell(new Phrase("JUNCTION BOX SIZE",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("10.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TOTAL TANK WEIGHT",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			PdfPCell Acessories = new PdfPCell(new Phrase("ACCESSORIES",title1));
			//cell.BackgroundColor = new iTextSharp.text.BaseColor(0, 150, 0);
			Acessories.Colspan = 4;
			Acessories.HorizontalAlignment = 1;
			table.AddCell(Acessories);

			table.AddCell(new PdfPCell(new Phrase("1.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TAP CHANGER",tab)));
			table.AddCell(new PdfPCell(new Phrase("AMPS",tab)));
			table.AddCell(new PdfPCell(new Phrase("POSITION",tab)));

			table.AddCell(new PdfPCell(new Phrase("2.",heading)));
			table.AddCell(new PdfPCell(new Phrase("WWTL",tab)));
			table.AddCell(new PdfPCell(new Phrase("OTI",tab)));
			table.AddCell(new PdfPCell(new Phrase("MARSHAL BOX",tab)));

			table.AddCell(new PdfPCell(new Phrase("3.",heading)));
			table.AddCell(new PdfPCell(new Phrase("PRV",tab)));
			table.AddCell(new PdfPCell(new Phrase("SUCHHOLZ RELAY",tab)));
			table.AddCell(new PdfPCell(new Phrase("MOG",tab)));

			table.AddCell(new PdfPCell(new Phrase("4.",heading)));
			table.AddCell(new PdfPCell(new Phrase("CONNECTING ROD SIZE",tab)));
			table.AddCell(new PdfPCell(new Phrase("LT",tab)));
			table.AddCell(new PdfPCell(new Phrase("HT",tab)));

			table.AddCell(new PdfPCell(new Phrase("5.",heading)));
			table.AddCell(new PdfPCell(new Phrase("OIL QUANTITY Ltr.",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			table.AddCell(new PdfPCell(new Phrase("6.",heading)));
			table.AddCell(new PdfPCell(new Phrase("TOTAL WEIGHT OF DTR",tab)));
			table.AddCell(new PdfPCell(new Phrase()));
			table.AddCell(new PdfPCell(new Phrase()));

			doc.Add(table);

			doc.Close();
			System.Diagnostics.Process.Start("C:\\Users\\Sid\\source\\repos\\Indian Power\\Indian Power\\bin\\Debug\\Final Result PDF.pdf");
		}

		private void Form1_Activated(object sender, EventArgs e)
		{
			text_Job_No.Focus();
		}

		private void Button2_Click_1(object sender, EventArgs e)
		{
			MessageBox
				.Show(Flux_Density.ToString());
		}

		private void Combo_Validated(object sender, EventArgs e)
		{
			text_capacity.Focus();
		}

		private void Text_Job_No_KeyPress(object sender, KeyPressEventArgs e)
		{
			if ((e.KeyChar >= 65 && e.KeyChar <= 91) || (e.KeyChar >= 97 && e.KeyChar <= 123))
			{
				e.Handled = true;
			}
		}

		private void Text_Total_Turns_Single_Disc_RO_TextChanged(object sender, EventArgs e)
		{
		}

		private void ComboBox5_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void Combo_Tapping_Step_Value_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (combo_Tapping_Step_Value.SelectedItem == null)
			{
				Tapping_Per_Step = 1.25;
				return;
			}
			Tapping_Per_Step = Double.Parse(combo_Tapping_Step_Value.SelectedItem.ToString());
			text_Tapping_Per_Step.Text = (Convert.ToDouble(Tapping_Per_Step)).ToString("0.00");
		}

		private void Button1_Click_1(object sender, EventArgs e)
		{
			new Form2().Show();
			this.Hide();
		}

		private void TextBox12_TextChanged(object sender, EventArgs e)
		{
			Tapping_Per_Step = Double.Parse(text_Tapping_Per_Step.Text);
			
			Total_Turns_Of_Coil = HV_Total_V * Final_Turn_per_volt;
			text_Turns_Other_Coils.Text = (Convert.ToDouble(Total_Turns_Of_Coil)).ToString("0.00");
		}

		private void Label150_Click(object sender, EventArgs e)
		{

		}

		private void TextBox8_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox6_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_capacity_KeyPress(object sender, KeyPressEventArgs e)
		{

		}

		private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			Ref_Standard = combo_Ref_Standard.SelectedItem.ToString();

			if (Ref_Standard.Equals("1180 - L1"))
				Rcap = cap;
			else if (Ref_Standard.Equals("1180 - L2"))
				Rcap = cap + (cap * 0.25);
			else
				Rcap = cap + (cap * 0.40);
		}

		private void PictureBox7_Click(object sender, EventArgs e)
		{
			text_capacity.Text = "";
		}

		private void TextBox5_TextChanged(object sender, EventArgs e)
		{

		}

		private void Label116_Click(object sender, EventArgs e)
		{

		}

		private void Combo_Gap_Of_Coils_SelectedIndexChanged(object sender, EventArgs e)
		{
			Gap_of_Coils = int.Parse(combo_Gap_Of_Coils.SelectedItem.ToString());
			CL_Core = HV_Outter_Dia_RO + Gap_of_Coils;
			text_CL_Core.Text = (Convert.ToDouble(CL_Core)).ToString("0.00");
			text_CORE_Window_Height.Text = (Convert.ToDouble(Window_Height_Core)).ToString("0.00");
			text_CORE_Diameter.Text = (Convert.ToDouble(Core_Inner_Dia)).ToString("0.00");
			text_CORE_CL.Text = (Convert.ToDouble(CL_Core)).ToString("0.00");
			Channel_Size = (CL_Core * 3) + (2 * Gap_of_Coils);
			text_Channel_Size.Text = (Convert.ToDouble(Channel_Size)).ToString("0.00");

			Channel_H = int.Parse(combo_Channel_H.SelectedItem.ToString());
			Channel_W = int.Parse(combo_Channel_W.SelectedItem.ToString());
			Tank_Length = Channel_Size + 20;
			text_Tank_Length.Text = (Convert.ToDouble(Tank_Length)).ToString("0.00");
		}

		private void Combo_Gap_HV_Coil_Tank_SelectedIndexChanged(object sender, EventArgs e)
		{
			Gap_HV_Coil_Tank = int.Parse(combo_Gap_HV_Coil_Tank.SelectedItem.ToString());
			Foot_Plate = HV_Outter_Dia_RO + Gap_HV_Coil_Tank;
			text_Foot_Plate.Text = (Convert.ToDouble(Foot_Plate)).ToString("0.00");
			Tank_Width = HV_Outter_Dia_RO + Gap_HV_Coil_Tank;
			text_Tank_Width.Text = (Convert.ToDouble(Tank_Width)).ToString("0.00");
		}

		private void Combo_Foot_Plate_H_SelectedIndexChanged(object sender, EventArgs e)
		{
			Foot_Plate_H = int.Parse(combo_Foot_Plate_H.SelectedItem.ToString());
			Foot_Plate_W = int.Parse(combo_Foot_Plate_W.SelectedItem.ToString());
			Height_Transformer = (Core_Inner_Dia * 0.97 * 2) + Window_Height_Core + Foot_Plate_H;
			text_Height_Transformer.Text = (Convert.ToDouble(Height_Transformer)).ToString("0.00");
		}

		private void Label75_Click(object sender, EventArgs e)
		{

		}

		private void TextBox2_TextChanged(object sender, EventArgs e)
		{

		}

		private void TabControl2_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void TextBox49_TextChanged(object sender, EventArgs e)
		{

		}

		private void ComboBox17_SelectedIndexChanged(object sender, EventArgs e)
		{
			No_Of_Selection_HV_coils = int.Parse(combo_Selection_Of_HV_Coil.SelectedItem.ToString());
			text_No_HV_Coils.Text = (Convert.ToDouble(No_Of_Selection_HV_coils)).ToString("0.00");

			if (No_Of_Selection_HV_coils == 2)
				Packing_Of_HV_Coils = 65;
			else if (No_Of_Selection_HV_coils == 4)
				Packing_Of_HV_Coils = 83;
			else if (No_Of_Selection_HV_coils == 6)
				Packing_Of_HV_Coils = 107;
			else
				Packing_Of_HV_Coils = 126;
			text_Packing_Of_HV_Coils.Text = (Convert.ToDouble(Packing_Of_HV_Coils)).ToString("0.00");

			XL_Length_Of_HV_Window = (Window_Height_Core - Packing_Of_HV_Coils) / No_Of_Selection_HV_coils;
			text_XL_Length_Of_HV_Window.Text = (Convert.ToDouble(XL_Length_Of_HV_Window)).ToString("0.00");
		}

		private void TextBox51_TextChanged(object sender, EventArgs e)
		{

		}

		private void TabPage9_Click(object sender, EventArgs e)
		{

		}

		private void ComboBox21_SelectedIndexChanged(object sender, EventArgs e)
		{
			Gap_of_Coils = int.Parse(combo_Gap_Of_Coils.SelectedItem.ToString());
			CL_Core = HV_Outter_Dia_RO + Gap_of_Coils;
			text_CL_Core.Text = (Convert.ToDouble(CL_Core)).ToString("0.00");
			text_CORE_Window_Height.Text = (Convert.ToDouble(Window_Height_Core)).ToString("0.00");
			text_CORE_Diameter.Text = (Convert.ToDouble(Core_Inner_Dia)).ToString("0.00");
			text_CORE_CL.Text = (Convert.ToDouble(CL_Core)).ToString("0.00");
			Channel_Size = (CL_Core * 3) + (2 * Gap_of_Coils);
			text_Channel_Size.Text = (Convert.ToDouble(Channel_Size)).ToString("0.00");
			Gap_HV_Coil_Tank = int.Parse(combo_Gap_HV_Coil_Tank.SelectedItem.ToString());
			Foot_Plate = HV_Outter_Dia_RO + Gap_HV_Coil_Tank;
			text_Foot_Plate.Text = (Convert.ToDouble(Foot_Plate)).ToString("0.00");
			Channel_H = int.Parse(combo_Channel_H.SelectedItem.ToString());
			Channel_W = int.Parse(combo_Channel_W.SelectedItem.ToString());
			Foot_Plate_H = int.Parse(combo_Foot_Plate_H.SelectedItem.ToString());
			Foot_Plate_W = int.Parse(combo_Foot_Plate_W.SelectedItem.ToString());
			Height_Transformer = (Core_Inner_Dia * 0.97 * 2) + Window_Height_Core + Foot_Plate_H;
		}

		private void Text_CL_Core_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox50_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Total_Coil_Outter_Dia_RO_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Coil_Outter_Dia_TextChanged(object sender, EventArgs e)
		{

		}

		private void ComboBox20_SelectedIndexChanged(object sender, EventArgs e)
		{
			HV_Paper_Gap_Factor = int.Parse(combo_HV_Paper_Gap_Factor.SelectedItem.ToString());
			HV_Outter_Dia = (Total_Layers_Single_Disc_RO * 2 * HV_Conductor_Size) + HV_Coil_Inner_Dia + HV_Paper_Gap_Factor;
			text_HV_Outter_Dia.Text = (Convert.ToDouble(HV_Outter_Dia)).ToString("0.00");
			HV_Outter_Dia_RO = Math.Round(HV_Outter_Dia);
			text_HV_Outter_Dia_RO.Text = (Convert.ToDouble(HV_Outter_Dia_RO)).ToString("0.00");
		}

		private void ComboBox18_SelectedIndexChanged(object sender, EventArgs e)
		{
			HV_Total_V = inV + ((inV * Tapping_Range_Max) / 100);
			text_HV_Total_V.Text = (Convert.ToDouble(HV_Total_V)).ToString("0.00");

			Total_HV_Turns = HV_Total_V * Final_Turn_per_volt;
			text_Total_HV_Turns.Text = (Convert.ToDouble(Total_HV_Turns)).ToString("0.00");

			Total_Turns_Single_Disc = Total_HV_Turns / No_Of_Selection_HV_coils;
			text_Total_Turns_Single_Disc.Text = (Convert.ToDouble(Total_Turns_Single_Disc)).ToString("0.00");

			Total_Turns_Single_Disc_RO = Math.Round(Total_Turns_Single_Disc);
			text_Total_Turns_Single_Disc_RO.Text = (Convert.ToDouble(Total_Turns_Single_Disc_RO)).ToString("0.00");

			text_Turns_Other_Coils.Text = (Convert.ToDouble(Total_Turns_Single_Disc_RO)).ToString("0.00");

			Turns_Per_Tapping = Math.Round((inV * Tapping_Per_Step * Final_Turn_per_volt) / 100);

			if (Tapping_Range_Max > 0 && Tapping_Range_Max <= 5)
			{
				Turns_Mid_C1 = Total_Turns_Single_Disc_RO;
				Turns_Mid_C2 = Total_Turns_Single_Disc_RO - Turns_Per_Tapping;
				Turns_Mid_C3 = Total_Turns_Single_Disc_RO - (2*Turns_Per_Tapping);
				text_Turns_Mid_C1.Text = (Convert.ToDouble(Turns_Mid_C1)).ToString("0.00");
				text_Turns_Mid_C2.Text = (Convert.ToDouble(Turns_Mid_C2)).ToString("0.00");
				text_Turns_Mid_C3.Text = (Convert.ToDouble(Turns_Mid_C3)).ToString("0.00");
			}

			else if (Tapping_Range_Max > 5 && Tapping_Range_Max <= 7.5)
			{
				Turns_Mid_C1 = Total_Turns_Single_Disc_RO;
				Turns_Mid_C2 = Total_Turns_Single_Disc_RO - Turns_Per_Tapping;
				Turns_Mid_C3 = Total_Turns_Single_Disc_RO - (2 * Turns_Per_Tapping);
				Turns_Mid_C4 = Total_Turns_Single_Disc_RO - (3 * Turns_Per_Tapping);
				text_Turns_Mid_C1.Text = (Convert.ToDouble(Turns_Mid_C1)).ToString("0.00");
				text_Turns_Mid_C2.Text = (Convert.ToDouble(Turns_Mid_C2)).ToString("0.00");
				text_Turns_Mid_C3.Text = (Convert.ToDouble(Turns_Mid_C3)).ToString("0.00");
				text_Turns_Mid_C4.Text = (Convert.ToDouble(Turns_Mid_C4)).ToString("0.00");
			}
			else
			{
				Turns_Mid_C1 = Total_Turns_Single_Disc_RO;
				Turns_Mid_C2 = Total_Turns_Single_Disc_RO - Turns_Per_Tapping;
				Turns_Mid_C3 = Total_Turns_Single_Disc_RO - (2 * Turns_Per_Tapping);
				Turns_Mid_C4 = Total_Turns_Single_Disc_RO - (3 * Turns_Per_Tapping);
				Turns_Mid_C5 = Total_Turns_Single_Disc_RO - (4 * Turns_Per_Tapping);
				text_Turns_Mid_C1.Text = (Convert.ToDouble(Turns_Mid_C1)).ToString("0.00");
				text_Turns_Mid_C2.Text = (Convert.ToDouble(Turns_Mid_C2)).ToString("0.00");
				text_Turns_Mid_C3.Text = (Convert.ToDouble(Turns_Mid_C3)).ToString("0.00");
				text_Turns_Mid_C4.Text = (Convert.ToDouble(Turns_Mid_C4)).ToString("0.00");
				text_Turns_Mid_C5.Text = (Convert.ToDouble(Turns_Mid_C5)).ToString("0.00");
			}

			Final_Total_Turns = Total_Turns_Single_Disc_RO * No_Of_Selection_HV_coils;
			text_Final_Total_Turns.Text = (Convert.ToDouble(Final_Total_Turns)).ToString("0.00");

			Turn_Per_Layer_Single_Disc = (XL_Length_Of_HV_Window / HV_Conductor_To_Use_Result) - HV_Conductor_To_Use_Result;
			text_Turn_Per_Layer_Single_Disc.Text = (Convert.ToDouble(Turn_Per_Layer_Single_Disc)).ToString("0.00");

			Total_Layers_Single_Disc = Total_Turns_Single_Disc / Turn_Per_Layer_Single_Disc;
			text_Total_Layers_Single_Disc.Text = (Convert.ToDouble(Total_Layers_Single_Disc)).ToString("0.00");

			Total_Layers_Single_Disc_RO = Math.Ceiling(Total_Layers_Single_Disc);
			text_Total_Layers_Single_Disc_RO.Text = (Convert.ToDouble(Total_Layers_Single_Disc_RO)).ToString("0.00");
		}

		private void TextBox54_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Tapping_Range_Max_TextChanged(object sender, EventArgs e)
		{
			if (text_Tapping_Range_Max.Text.Trim() == string.Empty)
			{
				text_Tapping_Range_Max.Text = "";
				return;
			}
			else
			{
				Tapping_Range_Max = Double.Parse(text_Tapping_Range_Max.Text);
				text_Tapping_Per_Max.Text = (Convert.ToDouble(Tapping_Range_Max)).ToString("0.00");
			}
		}

		private void TextBox55_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Tapping_Range_Min_TextChanged(object sender, EventArgs e)
		{
			if (text_Tapping_Range_Min.Text.Trim() == string.Empty)
			{
				text_Tapping_Range_Min.Text = "";
				return;
			}
			else
			{
				Tapping_Range_Min = Double.Parse(text_Tapping_Range_Min.Text);
				text_Tapping_Per_Min.Text = (Convert.ToDouble(Tapping_Range_Min)).ToString("0.00");
			}
		}

		private void TextBox57_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox53_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox52_TextChanged(object sender, EventArgs e)
		{

		}

		private void PictureBox4_MouseClick(object sender, MouseEventArgs e)
		{
			//Application.Exit(null);
		}

		private void Combo_Oil_Duct_Size_SelectedIndexChanged(object sender, EventArgs e)
		{
			Oil_Duct_Size = int.Parse(combo_Oil_Duct_Size.SelectedItem.ToString());
			HV_Coil_Inner_Dia = Total_Coil_Outter_Dia_RO + Oil_Duct_Size;
			text_HV_Coil_Inner_Dia.Text = (Convert.ToDouble(HV_Coil_Inner_Dia)).ToString("0.00");
		}

		private void Text_LV_Conductor_To_Use_Result_W_TextChanged(object sender, EventArgs e)
		{

		}

		private void Combo_Window_Padding_L_SelectedIndexChanged(object sender, EventArgs e)
		{
			Window_Padding_L = Double.Parse(combo_Window_Padding_L.SelectedItem.ToString());
			Window_Padding_U = Double.Parse(combo_Window_Padding_U.SelectedItem.ToString());
			Window_Height_Core = Math.Round(Total_Length_Coil) + Window_Padding_L + Window_Padding_U;
			text_Window_Height_Core.Text = (Convert.ToDouble(Window_Height_Core)).ToString("0.00");
		}

		private void CheckBox1_CheckedChanged(object sender, EventArgs e)
		{
			if(checkBox1.Checked)
			{
				Length_Of_Winding = Length_Of_One_Turn * Turns_Per_Layer + (2*Length_Of_One_Turn) ;
				text_Length_Of_Winding.Text = (Convert.ToDouble(Length_Of_Winding)).ToString("0.00");
				checkBox2.Checked = false;
			}
		}

		private void CheckBox2_CheckedChanged(object sender, EventArgs e)
		{
			if (checkBox2.Checked)
			{
				Length_Of_Winding = Length_Of_One_Turn * Turns_Per_Layer + Length_Of_One_Turn;
				text_Length_Of_Winding.Text = (Convert.ToDouble(Length_Of_Winding)).ToString("0.00");
				checkBox1.Checked = false;
			}
		}

		private void ComboBox14_SelectedIndexChanged(object sender, EventArgs e)
		{
			LV_Paper_Gap_Factor = int.Parse(combo_LV_Paper_Gap_Factor.SelectedItem.ToString());
			Total_Outter_Coil_Dia = Outter_Coil_Dia + LV_Paper_Gap_Factor + LV_Coil_Inner_Dia;
			text_Total_Coil_Outter_Dia.Text = (Convert.ToDouble(Total_Outter_Coil_Dia)).ToString("0.00");
		}

		private void ComboBox13_SelectedIndexChanged(object sender, EventArgs e)
		{
			Gap_Core_LV = int.Parse(combo_Gap_Core_LV.SelectedItem.ToString());
			LV_Coil_Inner_Dia  = Core_Inner_Dia + Gap_Core_LV;
			text_LV_Coil_Inner_Dia.Text = (Convert.ToDouble(LV_Coil_Inner_Dia)).ToString("0.00");
		}

		private void Text_Core_Inner_Dia_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Length_Of_Winding_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Chip_Size_L_TextChanged(object sender, EventArgs e)
		{
			if (text_Chip_Size_U.Text.Trim() == string.Empty || text_Chip_Size_L.Text.Trim() == string.Empty)
			{
				MessageBox.Show("Enter Chip Size Upper");
				return;
			}
			else
			{
				Chip_Size_U = Double.Parse(text_Chip_Size_U.Text);
				Chip_Size_L = Double.Parse(text_Chip_Size_L.Text);
				Total_Length_Coil = Total_Length_Of_Winding + Chip_Size_L + Chip_Size_U;
				text_Total_Length_Of_Coil.Text = (Convert.ToDouble(Total_Length_Coil)).ToString("0.00");
			}
		}

		private void Text_Total_Length_Coil_TextChanged(object sender, EventArgs e)
		{


		}

		private void Text_Length_Of_LV_Winding_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox37_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_No_Of_Conductor_W_TextChanged(object sender, EventArgs e)
		{
			if (text_No_Of_Conductor_W.Text.Trim() == string.Empty)
			{
				text_No_Of_Conductor_W.Text = "";
				return;
			}
			else
			{
				No_Of_Conductors_W = Double.Parse(text_No_Of_Conductor_W.Text);

				text_No_Of_Conductor_W.Text = (Convert.ToDouble(No_Of_Conductors_W)).ToString("0.00");
				text_No_Of_Cond_W.Text = (Convert.ToDouble(No_Of_Conductors_W)).ToString("0.00");

				Length_Of_One_Turn = LV_Conductor_To_Use_Result_W * No_Of_Conductors_W;
				text_Length_Of_One_Turn.Text = (Convert.ToDouble(Length_Of_One_Turn)).ToString("0.00");
			}
		}

		private void Text_No_Of_Conductor_D_TextChanged(object sender, EventArgs e)
		{
			if (text_No_Of_Conductor_D.Text.Trim() == string.Empty)
			{
				text_No_Of_Conductor_D.Text = "";
				return;
			}
			else
			{
				No_Of_Conductors_D = Double.Parse(text_No_Of_Conductor_D.Text);
				text_No_Of_Conductor_D.Text = (Convert.ToDouble(No_Of_Conductors_D)).ToString("0.00");
				text_No_Of_Cond_D.Text = (Convert.ToDouble(No_Of_Conductors_D)).ToString("0.00");
			}
		}

		private void Text_No_Of_Layers_TextChanged(object sender, EventArgs e)
		{
			if (text_No_Of_Layers.Text.Trim() == string.Empty)
			{
				text_No_Of_Layers.Text = "";
				return;
			}
			else
			{
				No_Of_Layers = Double.Parse(text_No_Of_Layers.Text);
				Turns_Per_Layer = total_LT_Turns / No_Of_Layers;
				text_Turns_Per_Layer.Text = (Convert.ToDouble(Turns_Per_Layer)).ToString("0.0");
				Outter_Coil_Dia = LV_Conductor_To_Use_Result_D * No_Of_Conductors_D * No_Of_Layers * 2;
				text_Coil_Outter_Dia.Text = (Convert.ToDouble(Outter_Coil_Dia)).ToString("0.000");
			}
		}

		private void Text_Final_Turns_per_Volt_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox30_TextChanged(object sender, EventArgs e)
		{
			if (text_HV_Insulation_Covering.Text.Trim() == string.Empty)
			{
				text_HV_Insulation_Covering.Text = "";
				return;
			}
			else
			{
				HV_Insulation_Covering = Double.Parse(text_HV_Insulation_Covering.Text);

				HV_Conductor_To_Use_Result = HV_Insulation_Covering + HV_Conductor_Dia;
				text_HV_Conductor_To_Use_Result.Text = (Convert.ToDouble(HV_Conductor_To_Use_Result)).ToString("0.00");

				HV_Conductor_Size = HV_Conductor_To_Use_Result;
				text_HV_Conductor_Size.Text = (Convert.ToDouble(HV_Conductor_Size)).ToString("0.00");
			}

		}

		private void GroupBox3_Enter(object sender, EventArgs e)
		{

		}

		private void SAVEToolStripMenuItem_Click(object sender, EventArgs e)
		{
			
		}

		private void Text_LV_Insulation_Covering_TextChanged(object sender, EventArgs e)
		{
			if (text_LV_Insulation_Covering.Text.Trim() == string.Empty)
			{
				text_LV_Insulation_Covering.Text = "";
				return;
			}
			else
			{
				LV_Insulation_Covering = Double.Parse(text_LV_Insulation_Covering.Text);

				LV_Conductor_To_Use_Result_D = LV_Conductor_Thickness + LV_Insulation_Covering;
				text_LV_Conductor_To_Use_Result_D.Text = (Convert.ToDouble(LV_Conductor_To_Use_Result_D)).ToString("0.00");

				LV_Conductor_To_Use_Result_W = LV_Conductor_Width + LV_Insulation_Covering;
				text_LV_Conductor_To_Use_Result_W.Text = (Convert.ToDouble(LV_Conductor_To_Use_Result_W)).ToString("0.00");

				total_LT_Turns = Final_Turn_per_volt * single_phase_v;
				text_Total_LT_Turns.Text = (Convert.ToDouble(total_LT_Turns)).ToString("0.00");

				text_Turn_Per_Volt_Next.Text = (Convert.ToDouble(Final_Turn_per_volt)).ToString("0.000");
			}
		}

		private void Text_capacity_KeyDown(object sender, KeyEventArgs e)
		{


		}

		private void Text_Conductor_Width_TextChanged(object sender, EventArgs e)
		{
			if (text_Conductor_Width.Text.Trim() == string.Empty)
			{
				text_Conductor_Width.Text = "";
				return; 
			}
			else
			{
				LV_Conductor_Width = Double.Parse(text_Conductor_Width.Text);

				LV_Conductor_Thickness = LV_Conductor_Size / LV_Conductor_Width;
				text_Conductor_Thickness.Text = (Convert.ToDouble(LV_Conductor_Thickness)).ToString("0.00");

				Bare_Conductor_Size_W = LV_Conductor_Width;
				text_Bare_Conductor_W.Text = (Convert.ToDouble(Bare_Conductor_Size_W)).ToString("0.0");

				Bare_Conductor_Size_D = LV_Conductor_Thickness;
				text_Bare_Conductor_D.Text = (Convert.ToDouble(Bare_Conductor_Size_D)).ToString("0.0");
			}

		}

		private void Text_HV_Turns_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_LT_Conductor_Area_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox24_TextChanged(object sender, EventArgs e)
		{
			if (text_No_of_Conductors.Text.Trim() == string.Empty)
			{
				text_No_of_Conductors.Text = "";
				return; 
			}
			else
			{
				No_of_Conductors = int.Parse(text_No_of_Conductors.Text);

				LV_Conductor_Size = LV_Conductor_Area / No_of_Conductors;
				text_LV_Conductor_Size.Text = (Convert.ToDouble(LV_Conductor_Size)).ToString("0.00");
			}
		}

		private void Text_Conductor_Density_TextChanged(object sender, EventArgs e)
		{
			if (text_Conductor_Density.Text.Trim() == string.Empty)
			{
				text_Conductor_Density.Text = "";
				return;
			}
			else
			{
				Conductor_Density = Double.Parse(text_Conductor_Density.Text);

				LV_Conductor_Area = outC / Conductor_Density;
				text_LV_Conductor_Area.Text = (Convert.ToDouble(LV_Conductor_Area)).ToString("0.00");

				HV_Conductor_Area = (inC / Conductor_Density) / Math.Sqrt(3);
				text_HV_Conductor_Area.Text = (Convert.ToDouble(HV_Conductor_Area)).ToString("0.00");

				HV_Conductor_Dia = 2 * (Math.Sqrt(HV_Conductor_Area / 3.14));
				text_HV_Conductor_Dia.Text = (Convert.ToDouble(HV_Conductor_Dia)).ToString("0.00");
			}
		}

		private void TabPage2_Click(object sender, EventArgs e)
		{

		}

		private void Label123_Click(object sender, EventArgs e)
		{

		}

		private void Text_LV_Turns_TextChanged(object sender, EventArgs e)
		{

		}

		private void Text_Tuns_Per_Volt_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox14_TextChanged(object sender, EventArgs e)
		{
			if(comboBox_Flux_Density.SelectedItem == null)
			{
				Flux_Density = 1.5;
				text_Flux_Density.Text = (Convert.ToDouble(Flux_Density)).ToString("0.00");
			}
		}

		private void Text_Flux_Constant_TextChanged(object sender, EventArgs e)
		{
			if (text_Flux_Constant.Text.Trim() == string.Empty)
			{
				text_Flux_Constant.Text = "";
				return;
			}
			else
			{
				flux_constant = Double.Parse(text_Flux_Constant.Text);

				Turn_per_volt = flux_constant / ELA;
				text_Tuns_Per_Volt.Text = (Convert.ToDouble(Turn_per_volt)).ToString("0.00000");

				LV_turns = single_phase_v * Turn_per_volt;
				text_LV_Turns.Text = (Convert.ToDouble(LV_turns)).ToString("0.00");

				HV_turns = Turn_per_volt * ((inV+(inV*Tapping_Range_Max)/100));
				text_HV_Turns.Text = (Convert.ToDouble(HV_turns)).ToString("0.00");

				Final_LV_turns = Math.Round(LV_turns);
				text_Final_LV_Turns.Text = (Convert.ToDouble(Final_LV_turns)).ToString("0.0");

				Final_Turn_per_volt = Final_LV_turns / single_phase_v;
				text_Final_Turns_per_Volt.Text = (Convert.ToDouble(Final_Turn_per_volt)).ToString("0.00000");

				Final_HV_turns = ((inV + (inV * Tapping_Range_Max)/100)) * Final_Turn_per_volt;
				text_Final_HV_Turns.Text = (Convert.ToDouble(Final_HV_turns)).ToString("0.0");

				text_Turn_Per_Volt_Next.Text = (Convert.ToDouble(Final_Turn_per_volt)).ToString("0.0");
			}
		}

		private void ComboBox_Flux_Density_SelectedIndexChanged(object sender, EventArgs e)
		{
			Flux_Density = Double.Parse(comboBox_Flux_Density.SelectedItem.ToString());
			text_Flux_Density.Text = (Convert.ToDouble(Flux_Density)).ToString("0.00");
		}

		private void EXITToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Application.Exit(null);
		}

		private void Text_capacity_TextChanged(object sender, EventArgs e)
		{
			String val = text_capacity.Text;
			if (text_capacity.Text.Trim() == string.Empty)
			{
				text_capacity.Text = "";
				//MessageBox.Show("Enter Capacity");
				return;
			}
			else
			{

				if (!char.IsDigit(val[val.Length-1]))
				{
					//MessageBox.Show("true");
					if (text_capacity.Text.Length == 1)
					{
						text_capacity.Text = "";
					}
					if (text_capacity.Text.Length > 0)
					{
						text_capacity.Text = text_capacity.Text.Substring(0, text_capacity.Text.Length - 2);
						text_capacity.Focus();
					}
					return;
				}
				//--------------------------
				cap = Double.Parse(text_capacity.Text);

				if (cap > 0 && cap <= 500)
					impedance = 4.5;
				else if (cap > 500 && cap <= 800)
					impedance = 5;
				else
					impedance = 6.25;

				text_Impedance.Text = (Convert.ToDouble(impedance)).ToString("0.00");


				Sel_CTA = (Math.Sqrt(Rcap * 1000)) / 5.9;
				text_Selection_Core_Area.Text = (Convert.ToDouble(Sel_CTA)).ToString("0.00");

				ELA = Sel_CTA / 3;
				text_Each_Limb_Area.Text = (Convert.ToDouble(ELA)).ToString("0.00");

				Dia_sqinch = 2 * (Math.Sqrt(ELA / 3.14));
				text_Diameter_SqInch.Text = (Convert.ToDouble(Dia_sqinch)).ToString("0.00");

				Dia_sqmm = Dia_sqinch * 25.4;
				text_Diameter_Sqmm.Text = (Convert.ToDouble(Dia_sqmm)).ToString("0.000");

				Core_Inner_Dia = Math.Round(Dia_sqmm);
				text_Core_Inner_Dia.Text = (Convert.ToDouble(Core_Inner_Dia)).ToString("0.00");

				text_Flux_Density.Text = comboBox_Flux_Density.SelectedIndex.ToString();

			}
		}

		private void Text_outA_TextChanged(object sender, EventArgs e)
		{

		}

		private void TextBox16_TextChanged(object sender, EventArgs e)
		{

		}

		private void SAVEASToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (printDialog1.ShowDialog() == DialogResult.OK)
			{
				printPreviewDialog1.ShowDialog();
			}
		}
		private void Button3_Click(object sender, EventArgs e)
		{

		}

		private void Button1_Click(object sender, EventArgs e)
		{
		
		}

		private void Button2_Click(object sender, EventArgs e)
		{

		}
	}
}
