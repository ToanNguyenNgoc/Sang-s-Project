using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

using ZedGraph;
using AForge;
using AForge.Neuro;
using AForge.Neuro.Learning;
using AForge.Controls;
using System.Globalization;
using System.Runtime.InteropServices;

//


namespace TVMH
{
    public partial class Form1 : Form
    {
        private const int NFEATURES = 7; // số lượng đầu vào đầu  ra mặc định: <Sinh lý trẻ em>	<Nghệ thuật tạo hình>	<Toán cơ sở>	<Tiếng Việt>	<Âm nhạc>	<Tin học>	<output>
        private const int MAXLOOKBACK = 10; // số lượng lớn nhất của input vector
        private const int MAXFEATURES = NFEATURES * MAXLOOKBACK; // kích thướt lớn nhất của net input vector
        private int nFeatures = NFEATURES; // kích thướt mặc định của input vector
     
        private int nData = 0; // số lượng record đọc từ file
        private const int MAXSAMPLES = 200000; //số lượng record lớn nhất trong demo. chỉnh sửa khi cần
        private double[,] data = new double[NFEATURES, MAXSAMPLES]; // kích thướt ma trận data
        private double[,] data_GV = new double[NFEATURES, MAXSAMPLES]; // kích thướt ma trận data
        private double[] minValues = new double[MAXSAMPLES]; // min không gian vector train
        private double[] maxValues = new double[MAXSAMPLES]; // max không gian vector train
        private double[] minTestValues = new double[MAXSAMPLES]; // min không gian vector test
        private double[] maxTestValues = new double[MAXSAMPLES]; // max không gian vector test
        private int[] outputIndex = new int[MAXSAMPLES];
        private int[] outputTestIndex = new int[MAXSAMPLES];
        int[] allowedFeatureIndicesSW = new int[NFEATURES] { 3, 4, 5, 6, 7, 8, 9 }; //<Sinh lý trẻ em>	<Nghệ thuật tạo hình>	<Toán cơ sở>	<Tiếng Việt>	<Âm nhạc>	<Tin học>
       
        private string[] featureNames = new string[MAXFEATURES];
        private string rawInputFileName = "";
        private string workingDirectory = "";


        //tham số Network  
        private bool networkLoaded = false; // tham số duoc load lên từ file?
        private static ActivationNetwork   network = null;
    
        BackPropagationLearning teacher = null; // thuật toán lan truyền ngực
        private double learningRate = 0.001;// tốc độ học
        private double momentum = 0.0;
        private double sigmoidAlphaValue = 2.0;// 
        private int iterations = 150000;// Số lần lặp mặc định
        private bool needToStop = false; //cờ chạy/dừng

        // tham số các điểm kiểm thử
        public PointPairList listTarget = new PointPairList(); // list cần kiểm thử
        public PointPairList listPredicted = new PointPairList(); //giá trị dự đoán



        public Form1()
        {
            InitializeComponent();
            ttLookBack.SetToolTip(lbLookBack, "Number of successive 6-value TVMH samples to combine as net input vector");
            ttInputData.SetToolTip(gbInputData, "Step 1:  load data from Menu with File > Open Stockwatch Data File");
            txtSigmoidAlpha.Text = sigmoidAlphaValue.ToString();
            txtLearningRate.Text = learningRate.ToString();
            txtMomentum.Text = momentum.ToString();
            txtMaxIterations.Text = iterations.ToString();
            listBoxInputVectors.HorizontalScrollbar = true;
            listBoxTestVectors.HorizontalScrollbar = true;
            txtStatus.Text = "Load dữ liệu từ File";
            zedGraphControl1.GraphPane.CurveList.Clear(); // Thẻ vẻ các điểm kiểm thử
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.Title.Text = "Thực nghiệm và dự báo";
            myPane.XAxis.Title.Text = "Thực nghiệm";
            myPane.YAxis.Title.Text = "Giá trị";
        }

        private void loadData()
        {
            StreamReader reader = null;
            nData = 0;
            comboBoxOpenIndex.Items.Clear();
            comboBoxHighIndex.Items.Clear();
            comboBoxLowIndex.Items.Clear();
            comboBoxCloseIndex.Items.Clear();
            listBoxData.Items.Clear();
            try
                {
                    // open selected file
                    reader = File.OpenText(rawInputFileName);
                    string str = null;
                    char[] delimiterChars = { ',' };


                    if ((str = reader.ReadLine()) != null) // feature names
                    {
                        string[] xy = str.Split(delimiterChars);
                        int nT = xy.Length;
                        for (int j = 0; j < nT; j++)
                        {
                            comboBoxOpenIndex.Items.Add(xy[j]);
                            comboBoxHighIndex.Items.Add(xy[j]);
                            comboBoxLowIndex.Items.Add(xy[j]);
                            comboBoxCloseIndex.Items.Add(xy[j]);
                        }
                        comboBoxOpenIndex.SelectedIndex = allowedFeatureIndicesSW[0];
                        comboBoxHighIndex.SelectedIndex = allowedFeatureIndicesSW[1];
                        comboBoxLowIndex.SelectedIndex = allowedFeatureIndicesSW[2];
                        comboBoxCloseIndex.SelectedIndex = allowedFeatureIndicesSW[3];

                        for (int j = 0; j < nFeatures; j++) // feature names
                        {
                            featureNames[j] = xy[allowedFeatureIndicesSW[j]];
                            comboBoxTargets.Items.Add(featureNames[j]);
                        }
                        comboBoxTargets.Enabled = true;
                        comboBoxTargets.SelectedIndex = 3;
                    }
                    else
                    {
                        txtStatus.Text = "Couldn't read first line (feature/target names)";
                        return;
                    }
                    nData = 1;
                    while ((str = reader.ReadLine()) != null) // read values
                    {
                      
                        string[] xy = str.Split(delimiterChars);
                        string dt = xy[1];
                        for (int j = 0; j < nFeatures; j++) // feature values
                            data[j, nData] = double.Parse(xy[allowedFeatureIndicesSW[j]]);
                        string lbSTR = nData.ToString() + " " + dt+ " " + data[0,nData]+" "+data[1,nData]+" "+data[2,nData]+" "+data[3,nData]
                            + " " + data[4, nData] + " " + data[5, nData] + " " + data[6, nData]  ;
                        listBoxData.Items.Add(lbSTR);
                        nData++;

                    }
                    //listView1.View = View.Details;
                    ttDataListBox.SetToolTip(listBoxData, "Select range of samples to add to training or test set with Shift + Mouse click, then left click to select which set");
                    txtStatus.Text = "Read " + nData.ToString() + " TVMH values from " + Path.GetFileName(rawInputFileName);
                    for(int i=0;i<MAXLOOKBACK;i++)
                    {
                        comboBoxLookBack.Items.Add(i+1);
                    }
                    comboBoxLookBack.Enabled = true;
                    comboBoxLookBack.SelectedIndex = 0;
                    int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
                    int nNodes = selectedLookBack * NFEATURES * 2;
                    txtHiddenNodes.Text = nNodes.ToString();
                    comboBoxOpenIndex.SelectedIndex = allowedFeatureIndicesSW[0];
                }
                catch (Exception e7)
                {
                    MessageBox.Show("Failed reading the file\n" + e7.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    // close file
                    if (reader != null)
                        reader.Close();
                }
            

        }
        private void openStockwatchDataFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamReader reader = null;
               // read maximum MAXSAMPLES points
                rawInputFileName = openFileDialog1.FileName;
                try
                {
                    // open selected file
                    reader = File.OpenText(rawInputFileName);
                    string str = null;
                    char[] delimiterChars = { ',' };


                    if ((str = reader.ReadLine()) != null) // feature names
                    {
                        string[] xy = str.Split(delimiterChars);
                        int nT = xy.Length;
                        for (int j = 0; j < nT; j++)
                        {
                            comboBoxOpenIndex.Items.Add(xy[j]);
                            comboBoxHighIndex.Items.Add(xy[j]);
                            comboBoxLowIndex.Items.Add(xy[j]);
                            comboBoxCloseIndex.Items.Add(xy[j]);
                        }
                        comboBoxOpenIndex.SelectedIndex = allowedFeatureIndicesSW[0];
                        comboBoxHighIndex.SelectedIndex = allowedFeatureIndicesSW[1];
                        comboBoxLowIndex.SelectedIndex = allowedFeatureIndicesSW[2];
                        comboBoxCloseIndex.SelectedIndex = allowedFeatureIndicesSW[3];
                    }
                    txtStatus.Text = "Modify columns to identify OHLC if necessary and then click \"Load Data\" button";
                    
                    
                }
                catch (Exception e7)
                {
                    MessageBox.Show("Failed reading the file\n" + e7.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    // close file
                    if (reader != null)
                        reader.Close();
                }
            }
        }

 

        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            listBoxData.SelectedIndex = listBoxData.IndexFromPoint(e.Location);
            if (e.Button == MouseButtons.Right)
            {
                //select the item under the mouse pointer
                listBoxData.SelectedIndex = listBoxData.IndexFromPoint(e.Location);
                if (listBoxData.SelectedIndex != -1)
                {
                    cmDataBox.Show(listBoxData, new System.Drawing.Point(e.X, e.Y));
                }
            }

        }

        private void addToTraingSetToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void addToTraingSetToolStripMenuItem_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem;
            if (item.Text == "Add to Training Set")
            {
                txtStatus.Text = "Moving selected items to training set...please wait";
                Application.DoEvents();
                foreach (Object selectedItem in listBoxData.SelectedItems)
                {
                    listBoxTrainingSet.Items.Add(selectedItem);
                }
                for (int i = listBoxData.SelectedIndices.Count - 1; i >= 0; i--) // bad things happen if we try to remove items early in the list, so start at the end 
                {
                    listBoxData.Items.RemoveAt(listBoxData.SelectedIndices[i]);
                }
                sortListBox(listBoxData);
                sortListBox(listBoxTrainingSet);
                txtStatus.Text = "Selected items moved to training set";
            }
            else if (item.Text == "Add to Test Set")
            {
                txtStatus.Text = "Moving selected items to test set...please wait";
                Application.DoEvents();
                foreach (Object selectedItem in listBoxData.SelectedItems)
                {
                    listBoxTestSet.Items.Add(selectedItem);
                }
                for (int i = listBoxData.SelectedIndices.Count - 1; i >= 0; i--) // bad things happen if we try to remove items early in the list, so start at the end 
                {
                    listBoxData.Items.RemoveAt(listBoxData.SelectedIndices[i]);
                }
                sortListBox(listBoxData);
                sortListBox(listBoxTestSet);
                txtStatus.Text = "Selected items moved to test set";
            }

        }

        private void sortListBox(ListBox whichBox)
        {
            for (int i = 0; i < whichBox.Items.Count - 1;i++ )
            {
                string thisItem = whichBox.Items[i].ToString();
                int thisNumber = Convert.ToInt32(thisItem.Split(' ')[0]);
                for (int j = i + 1; j < whichBox.Items.Count; j++)
                {
                    thisItem = whichBox.Items[j].ToString();
                    int nextNumber = Convert.ToInt32(thisItem.Split(' ')[0]);
                    if (thisNumber > nextNumber)
                    {
                        Object tempItem = whichBox.Items[i];
                        whichBox.Items[i] = whichBox.Items[j];
                        whichBox.Items[j] = tempItem;
                        thisNumber = nextNumber;
                    }
                }
            }
        }

        private void listBoxTrainingSet_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            listBoxTrainingSet.SelectedIndex = listBoxTrainingSet.IndexFromPoint(e.Location);
            if (e.Button == MouseButtons.Right)
            {
                //select the item under the mouse pointer
                listBoxTrainingSet.SelectedIndex = listBoxTrainingSet.IndexFromPoint(e.Location);
                if (listBoxTrainingSet.SelectedIndex != -1)
                {
                    cmTrainingBox.Show(listBoxTrainingSet, new System.Drawing.Point(e.X, e.Y));
                }
            }

        }

        private void toolStripMenuItem1_DropDownItemClicked(object sender, System.Windows.Forms.ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem;
            if (item.Text == "Return selected to data set")
            {
                foreach (Object selectedItem in listBoxTrainingSet.SelectedItems)
                {
                    listBoxData.Items.Add(selectedItem);
                }
                for (int i = listBoxTrainingSet.SelectedIndices.Count - 1; i >= 0; i--) // bad things happen if we try to remove items early in the list, so start at the end 
                {
                    listBoxTrainingSet.Items.RemoveAt(listBoxTrainingSet.SelectedIndices[i]);
                }
                txtStatus.Text = "Returning training items to data set...please wait";
                Application.DoEvents();
                sortListBox(listBoxData);
                sortListBox(listBoxTrainingSet);
                txtStatus.Text = "Training items returned to data set";
            }
            else if (item.Text == "Move selected to test set")
            {
                foreach (Object selectedItem in listBoxTrainingSet.SelectedItems)
                {
                    listBoxTestSet.Items.Add(selectedItem);
                }
                for (int i = listBoxTrainingSet.SelectedIndices.Count - 1; i >= 0; i--) // bad things happen if we try to remove items early in the list, so start at the end 
                {
                    listBoxTrainingSet.Items.RemoveAt(listBoxTrainingSet.SelectedIndices[i]);
                }
                txtStatus.Text = "Moving training items to test set...please wait";
                Application.DoEvents();
                sortListBox(listBoxTestSet);
                sortListBox(listBoxTrainingSet);
                txtStatus.Text = "Training items moved to test set";
            }

        }

        private void toolStripMenuItem3_DropDownItemClicked(object sender, System.Windows.Forms.ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem;
            if (item.Text == "Return selected to data set")
            {
                foreach (Object selectedItem in listBoxTestSet.SelectedItems)
                {
                    listBoxData.Items.Add(selectedItem);
                }
                for (int i = listBoxTestSet.SelectedIndices.Count - 1; i >= 0; i--) // bad things happen if we try to remove items early in the list, so start at the end 
                {
                    listBoxTestSet.Items.RemoveAt(listBoxTestSet.SelectedIndices[i]);
                }
                txtStatus.Text = "Returning test items to data set...please wait";
                Application.DoEvents();
                sortListBox(listBoxData);
                sortListBox(listBoxTestSet);
                txtStatus.Text = "Test items returned to data set";
            }
            else if (item.Text == "Move selected to training set")
            {
                foreach (Object selectedItem in listBoxTestSet.SelectedItems)
                {
                    listBoxTrainingSet.Items.Add(selectedItem);
                }
                for (int i = listBoxTestSet.SelectedIndices.Count - 1; i >= 0; i--) // bad things happen if we try to remove items early in the list, so start at the end 
                {
                    listBoxTestSet.Items.RemoveAt(listBoxTestSet.SelectedIndices[i]);
                }
                txtStatus.Text = "Moving test items to training set...please wait";
                sortListBox(listBoxTestSet);
                sortListBox(listBoxTrainingSet);
                txtStatus.Text = "Test items moved to training set";
            }
        }

        private void listBoxTestSet_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            listBoxTestSet.SelectedIndex = listBoxTestSet.IndexFromPoint(e.Location);
            if (e.Button == MouseButtons.Right)
            {
                //select the item under the mouse pointer
                listBoxTestSet.SelectedIndex = listBoxTestSet.IndexFromPoint(e.Location);
                if (listBoxTestSet.SelectedIndex != -1)
                {
                    cmTestSet.Show(listBoxTestSet, new System.Drawing.Point(e.X, e.Y));
                }
            }

        }

        private void comboBoxLookBack_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
            int nNodes = selectedLookBack * NFEATURES * 2;
            txtHiddenNodes.Text = nNodes.ToString();
        }
        double errorToancuc = 0D;
        private void trainNetwork()
        {
            button5.Enabled = true;
            button2.Enabled = true;
            int nInputVectors = listBoxTrainingSet.Items.Count;
            if (nInputVectors == 0)
            {
                MessageBox.Show("Please move some items from the data set to the training set in the Data tab");
                return;
            }
            double[][] input = new double[nInputVectors][];
            double[][] output = new double[nInputVectors][];

            int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
            int thisNFeature = selectedLookBack * NFEATURES;
            int thisHiddenNodes = Convert.ToInt32(txtHiddenNodes.Text);// số noron lớp ẩn
            if (!networkLoaded)// Load mạng nếu có lưu
            {
                network = new AForge.Neuro.ActivationNetwork(
                    new AForge.Neuro.BipolarSigmoidFunction(sigmoidAlphaValue),
                    thisNFeature, thisHiddenNodes, 1);
            }// Không thì làm mới 

            // tạo mạng học mới teacher
            teacher = new AForge.Neuro.Learning.BackPropagationLearning(network);

            teacher.LearningRate = learningRate;
            teacher.Momentum = momentum;

            // / tạo danh sách dữ liệu huấn luyện in put và output
            for (int i = 0; i < nInputVectors; i++)
            {
                input[i] = new double[thisNFeature];
                output[i] = new double[1];
            }
            // lấy dữ liệu đưa vào input
            string str = null;
            char[] delimiterChars = { ' ' };
            int nVectors = 0;
            int n = 0;
            for (int i = 0; i < nInputVectors - thisNFeature - 1; i++)
            {
                str = listBoxTrainingSet.Items[i].ToString();
                string[] xy = str.Split(delimiterChars);
                int thisInd = Convert.ToInt32(xy[0]); // index into data array
                n = 0;
                for (int j = 0; j < selectedLookBack; j++)
                {
                    for (int k = 0; k < NFEATURES - 1; k++)
                    {
                        input[i][n] = data[k, thisInd + j];
                        n++;
                    }
                }

                int outInd = Convert.ToInt32(comboBoxTargets.SelectedIndex);
                output[i][0] = data[outInd, thisInd + selectedLookBack - 1];
                outputIndex[i] = thisInd + selectedLookBack - 1;
                nVectors++;
            }
            scaleIO(nVectors, thisNFeature, input, output);

            listInputVectors(nVectors, thisNFeature, input);

            listOutputVectors(nVectors, output);

            int nIter = 0;
            double[] networkInput = new double[thisNFeature];
            double learningError = 0D;
            int viewPrediction = 500;// 500 lần lặp làm mới danh sách kiểm thử 1 lần
            while (!needToStop)
            {
                double error = teacher.RunEpoch(input, output) / nData;
                learningError = 0D;
                if (nIter % viewPrediction == 0)
                    listBoxPrediction.Items.Clear();
                for (int i = 0; i < nVectors; i++)
                {
                    // chuyển input 2 chiều thành 1 chiều chuẩn bị huấn luyện
                    for (int j = 0; j < thisNFeature; j++)
                    {
                        networkInput[j] = input[i][j];
                    }
                    double learningErrorTmp;// Sai số học
                    double[] solN = network.Compute(networkInput); // + 0.85) / factor + yMin;
                    learningErrorTmp = Math.Abs(solN[0] - output[i][0]);
                    learningError = learningError + learningErrorTmp;
                    if (nIter % viewPrediction == 0)
                    {
                        double tmin = minValues[i];
                        double diff = maxValues[i] - tmin;
                        double iv = (solN[0] + .9D) / 1.8D * diff + tmin;
                        string[] p = listBoxTargets.Items[i].ToString().Split(' ');
                        listBoxPrediction.Items.Add(Math.Abs((double.Parse(p[1]) - learningErrorTmp)).ToString() + "      " + iv.ToString());
                    }

                }
                error = learningError / nData;
                if (nIter % viewPrediction == 0)
                {
                    listBoxPrediction.SelectedIndex = listBoxTargets.SelectedIndex;
                    Application.DoEvents();
                }

                nIter++;
                if (nIter >= iterations)
                    needToStop = true;
                txtStatus.Text = "Lần lặp " + nIter.ToString() + " sai số " + error.ToString();
                errorToancuc = error;
                Application.DoEvents();
            }
            needToStop = false;
            button2.Enabled = false;
            

        }

        private void listInputVectors(int nv, int n, double[][] input)
        {
            listBoxInputVectors.Items.Clear();
            for (int i = 0; i < nv; i++)
            {
                string strI = listBoxTrainingSet.Items[i].ToString();
                string[] xy = strI.Split(' ');
                string str = xy[0];
                double tmin = minValues[i];
                double diff = maxValues[i] - tmin;
                for (int j = 0; j < n; j++)
                {
                    double iv = (input[i][j] + .9D) / 1.8D * diff + tmin;
                    str += " " + iv.ToString("##.##");
                }
                listBoxInputVectors.Items.Add(str);
            }
        }

        private void listOutputVectors(int nv, double[][] output)
        {
            listBoxTargets.Items.Clear();
            for (int i = 0; i < nv; i++)
            {
                double tmin = minValues[i];
                double diff = maxValues[i] - tmin;
                double iv = (output[i][0] + .9D) / 1.8D * diff + tmin;
                string str = outputIndex[i].ToString() + " " + iv.ToString("##.##");
                listBoxTargets.Items.Add(str);
            }
        }

        private void scaleIO(int nv, int ni, double[][] input, double[][] output)
        {
            for (int i = 0; i < nv; i++)
            {
                double tmin = input[i][0];
                double tmax = tmin;
                for (int j = 0; j < ni; j++)
                {
                    tmin = Math.Min(tmin, input[i][j]);
                    tmax = Math.Max(tmax, input[i][j]);
                }
                tmin = Math.Min(tmin, output[i][0]);
                tmax = Math.Max(tmax, output[i][0]);
                minValues[i] = tmin;
                maxValues[i] = tmax;
                double diff = tmax - tmin;
                for (int j = 0; j < ni; j++)
                {
                    input[i][j] = (input[i][j] - tmin) / diff * 1.8 - 0.9D;
                }
                output[i][0] = (output[i][0] - tmin) / diff * 1.8 - 0.9D;
            }
        }
    

        private void button1_Click(object sender, EventArgs e)
        {
            iterations = Convert.ToInt32(txtMaxIterations.Text);
        }

        private void btnTrain_Click(object sender, EventArgs e)
        {
            trainNetwork();
            button8.Enabled = true; // Enable "Resume Training" button
            button4.Enabled = true; // Enable "Save Network" button
            button7.Enabled = true; // Enable "Save Vectors" button on Training tab
            btnTest.Enabled = true; // Enable "Start Test" button on Test tab;


        }

        private void button2_Click(object sender, EventArgs e)
        {
            needToStop = true;
        }

        private void listBoxTargets_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBoxPrediction.SelectedIndex = listBoxTargets.SelectedIndex;
            listBoxInputVectors.SelectedIndex = listBoxTargets.SelectedIndex;
            Application.DoEvents();
        }

        private void listBoxInputVectors_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBoxPrediction.SelectedIndex = listBoxInputVectors.SelectedIndex;
            listBoxTargets.SelectedIndex = listBoxInputVectors.SelectedIndex;

        }

        private void testNetwork()
        {
            listTarget.Clear();
            listPredicted.Clear();
            int nInputVectors = listBoxTestSet.Items.Count;
            if (nInputVectors == 0)
            {
                MessageBox.Show("Vui lòng chọn items từ tập huấn luyên sang tập kiểm tra trong tab xử lý dữ liệu");
                return;
            }
            double[][] input = new double[nInputVectors][];
            double[][] output = new double[nInputVectors][];

            int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
            int thisNFeature = selectedLookBack * NFEATURES;

            // create arrays to hold training data and output result
            for (int i = 0; i < nInputVectors; i++)
            {
                input[i] = new double[thisNFeature];
                output[i] = new double[1];
            }

            // populate the input vector with items from the test list taken LookBack at a time
            string str = null;
            char[] delimiterChars = { ' ' };
            int nVectors = 0;
            int n = 0;
            for (int i = 0; i < nInputVectors; i++)//for (int i = 0; i < nInputVectors - thisNFeature - 1; i++)
            {
                str = listBoxTestSet.Items[i].ToString();
                string[] xy = str.Split(delimiterChars);
                int thisInd = Convert.ToInt32(xy[0]); // index into data array
                n = 0;
                for (int j = 0; j < selectedLookBack; j++)
                {
                    for (int k = 0; k < NFEATURES - 1; k++)
                    {
                        input[i][n] = data[k, thisInd + j];
                        n++;
                    }
                }
                //  MessageBox.Show(comboBoxTargets.SelectedItem.ToString());
                int outInd = Convert.ToInt32(comboBoxTargets.SelectedIndex);
                output[i][0] = data[outInd, thisInd + selectedLookBack - 1];
                outputTestIndex[i] = thisInd + selectedLookBack;
                nVectors++;
            }
            scaleTestIO(nVectors, thisNFeature, input, output);
           listTestInputVectors(nVectors, thisNFeature, input);
           listTestOutputVectors(nVectors, output);
            double[] networkInput = new double[thisNFeature];
            double learningError = 0D;
            // run epoch of learning procedure
            learningError = 0D;
            listBoxTestPrediction.Items.Clear();
            listPredicted.Clear();
            for (int i = 0; i < nVectors; i++)
            {
                // put values from current window as network's input
                for (int j = 0; j < thisNFeature; j++)
                {
                    networkInput[j] = input[i][j];
                }
                double learningErrorTmp = 0D;

                double[] solN = network.Compute(networkInput); // + 0.85) / factor + yMin;
                
                learningErrorTmp = Math.Abs(solN[0] - output[i][0]);
                learningError = learningError + learningErrorTmp;
                double tmin = minTestValues[i];
                double diff = maxTestValues[i] - tmin;
                double iv = (solN[0] + .9D) / 1.8D * diff + tmin;
                /// double iv = solN[0];

                string[] p = listBoxTestTargets.Items[i].ToString().Split(' ');

                //    double aa = double.Parse(p[1].ToString()) - learningErrorTmp;
                double aa = iv - learningErrorTmp;
                //   double aa = iv - errorToancuc;
                listBoxTestPrediction.Items.Add(outputTestIndex[i].ToString() + " " + iv.ToString("#.##"));
                // listBoxTestPrediction.Items.Add(  " " + iv.ToString("#.##"));
                listPredicted.Add((double)i, iv);

            }
            txtStatus.Text = "Sai số kiểm tra: " + (learningError/nVectors).ToString();
            zedGraphControl1.GraphPane.CurveList.Clear(); // Test tab graph of Actual vs Predicted
            GraphPane myPane = zedGraphControl1.GraphPane;
        
            LineItem targetCurve = myPane.AddCurve("Giá trị thực nghiệm",
                  listTarget, Color.Black, SymbolType.Circle); // plot centroid first to appear on top of data (plot order controls Z-order
            targetCurve.Line.IsVisible = false;
            targetCurve.Symbol.Size = 5.0F;
            targetCurve.Symbol.Fill = new Fill(Color.Black);

            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();

            LineItem predictedCurve = myPane.AddCurve("Giá trị dự báo",
                  listPredicted, Color.Red, SymbolType.Circle); // plot centroid first to appear on top of data (plot order controls Z-order
       
            targetCurve.Line.IsVisible = true;
            targetCurve.Symbol.Size = 5.0F;
            targetCurve.Symbol.Fill = new Fill(Color.Red);

            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();

        }

        private void listTestInputVectors(int nv, int n, double[][] input)
        {
            listBoxTestVectors.Items.Clear();
            for (int i = 0; i < nv; i++)
            {
                string strI = listBoxTestSet.Items[i].ToString();
                string[] xy = strI.Split(' ');
                string str = xy[0];
                double tmin = minTestValues[i];
                double diff = maxTestValues[i] - tmin;
                for (int j = 0; j < n; j++)
                {

                    double iv = (input[i][j] + .9D) / 1.8D * diff + tmin;
                    str += " " + iv.ToString("##.##");
                }
                listBoxTestVectors.Items.Add(str);
            }
        }
        private void listTestInputVectors_GV(int nv, int n, double[][] input)
        {
            lstKetquaGV.Items.Clear();
            for (int i = 0; i < nv; i++)
            {
                string strI = lstKetquaGV.Items[i].ToString();
                string[] xy = strI.Split(' ');
                string str = xy[0];
                double tmin = minTestValues[i];
                double diff = maxTestValues[i] - tmin;
                for (int j = 0; j < n; j++)
                {

                    double iv = (input[i][j] + .9D) / 1.8D * diff + tmin;
                    str += " " + iv.ToString("##.##");
                }
                lstKetquaGV.Items.Add(str);
            }
        }
        private void listTestOutputVectors(int nv, double[][] output)
        {
            listBoxTestTargets.Items.Clear();
            listTarget.Clear();
            for (int i = 0; i < nv; i++)
            {
                double tmin = minTestValues[i];
                double diff = maxTestValues[i] - tmin;
                double iv = (output[i][0] + .9D) / 1.8D * diff + tmin;
                listTarget.Add((double)i, iv);
                string str = outputTestIndex[i].ToString() + " " + iv.ToString("##.##");
                listBoxTestTargets.Items.Add(str);
            }
        }

        private void scaleTestIO(int nv, int ni, double[][] input, double[][] output)
        {
            for (int i = 0; i < nv; i++)
            {
                double tmin = input[i][0];
                double tmax = tmin;
                for (int j = 0; j < ni; j++)
                {
                    tmin = Math.Min(tmin, input[i][j]);
                    tmax = Math.Max(tmax, input[i][j]);
                }
                tmin = Math.Min(tmin, output[i][0]);
                tmax = Math.Max(tmax, output[i][0]);
                minTestValues[i] = tmin;
                maxTestValues[i] = tmax;
                double diff = tmax - tmin;
                for (int j = 0; j < ni; j++)
                {
                    input[i][j] = (input[i][j] - tmin) / diff * 1.8 - 0.9D;
                }
                output[i][0] = (output[i][0] - tmin) / diff * 1.8 - 0.9D;
            }
        }
        private void scaleTestIO_SV(int nv, int ni, double[][] input, double[][] output)
        {
            for (int i = 0; i < nv; i++)
            {
                double tmin = input[i][0];
                double tmax = tmin;
                for (int j = 0; j < ni; j++)
                {
                    tmin = Math.Min(tmin, input[i][j]);
                    tmax = Math.Max(tmax, input[i][j]);
                }
                tmin = Math.Min(tmin, output[i][0]);
                tmax = Math.Max(tmax, output[i][0]);
                minTestValues[i] = tmin;
                maxTestValues[i] = tmax;
                double diff = tmax - tmin;
                for (int j = 0; j < ni; j++)
                {
                    input[i][j] = (input[i][j] - tmin) / diff * 1.8 - 0.9D;
                }
                output[i][0] = (output[i][0] - tmin) / diff * 1.8 - 0.9D;
            }
        }
        private void scaleTestIO_GV(int nv, int ni, double[][] input, double[][] output)
        {
            for (int i = 0; i < nv; i++)
            {
                double tmin = input[i][0];
                double tmax = tmin;
                for (int j = 0; j < ni; j++)
                {
                    tmin = Math.Min(tmin, input[i][j]);
                    tmax = Math.Max(tmax, input[i][j]);
                }
                tmin = Math.Min(tmin, output[i][0]);
                tmax = Math.Max(tmax, output[i][0]);
                minTestValues[i] = tmin;
                maxTestValues[i] = tmax;
                double diff = tmax - tmin;
                for (int j = 0; j < ni; j++)
                {
                    input[i][j] = (input[i][j] - tmin) / diff * 1.8 - 0.9D;
                }
                output[i][0] = (output[i][0] - tmin) / diff * 1.8 - 0.9D;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                testNetwork();
                button6.Enabled = true; // enable "Save Vectors" button on Test tab
            }
            catch(Exception ex)
            {
                MessageBox.Show("Vui lòng nhập liệu");
            }
        }

        private void listBoxTestVectors_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBoxTestPrediction.SelectedIndex = listBoxTestVectors.SelectedIndex;
            listBoxTestTargets.SelectedIndex = listBoxTestVectors.SelectedIndex;

        }

        private void listBoxTestTargets_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBoxTestPrediction.SelectedIndex = listBoxTestVectors.SelectedIndex;
            listBoxTestTargets.SelectedIndex = listBoxTestVectors.SelectedIndex;

        }

        public string saveActivationNetworkWeights()
        {


            string outS = "";
            string retVal = "Save cancelled";
            saveFileDialog1.DefaultExt = "csv";
            saveFileDialog1.Filter =
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            if (workingDirectory == "")
                saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            else
                saveFileDialog1.InitialDirectory = workingDirectory;
            outS = comboBoxLookBack.SelectedIndex.ToString()+"_"+comboBoxTargets.SelectedIndex.ToString();
            saveFileDialog1.FileName = "Weights_" + outS + ".csv";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workingDirectory = Path.GetDirectoryName(saveFileDialog1.FileName);
                retVal = saveFileDialog1.FileName + ".csv saved";
                StreamWriter sw = new StreamWriter(saveFileDialog1.FileName + ".csv");
                sw.WriteLine(rawInputFileName);
                sw.WriteLine(outS);
                int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
                int thisNFeature = selectedLookBack * NFEATURES;
                outS = thisNFeature.ToString();
                sw.WriteLine(outS);

                outS = "";
                int nLayer = network.Layers.Length;
                sw.WriteLine(nLayer.ToString()); // write # of layers
                for (int i = 0; i < nLayer - 1; i++)
                {
                    outS += network.Layers[i].Neurons.Length.ToString() + ",";
                }
                outS += network.Layers[nLayer - 1].Neurons.Length.ToString();
                sw.WriteLine(outS); // write number of neurons for each layer;


                for (int i = 0; i < nLayer; i++)
                {
                    Layer layer = network.Layers[i];
                    int nNeurons = layer.Neurons.Length;
                    for (int j = 0; j < nNeurons; j++)
                    {

                        ActivationNeuron neuron = (ActivationNeuron)layer.Neurons[j];
                        double threshold = neuron.Threshold;
                        sw.WriteLine(threshold.ToString()); // write threshold for ith layer, jth neuron
                        int nWeights = neuron.Weights.Length;
                        outS = "";
                        if (nWeights == 1)
                        {
                            sw.WriteLine(neuron.Weights[0].ToString()); // write weight if only 1 neuron for nth layer
                        }
                        else
                        {
                            for (int k = 0; k < nWeights - 1; k++)
                            {
                                outS += neuron.Weights[k].ToString() + ",";
                            }
                            outS += neuron.Weights[nWeights - 1].ToString();
                            sw.WriteLine(outS); // write weights for nth layer
                        }
                    }

                }
                sw.Close();
            }
            return retVal;
        }

        public string loadActivationNetworkWeights()
        {

            string outS = "";
            string retVal = "Save cancelled";
            openFileDialog1.DefaultExt = "csv";
            openFileDialog1.Filter =
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            if (workingDirectory == "")
                openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            else
                openFileDialog1.InitialDirectory = workingDirectory;
            outS = comboBoxLookBack.SelectedIndex.ToString()+"_"+comboBoxTargets.SelectedIndex.ToString();
            openFileDialog1.FileName = "Weights_" + outS + ".csv";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workingDirectory = Path.GetDirectoryName(openFileDialog1.FileName);
                retVal = saveFileDialog1.FileName + ".csv saved";
                int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
                int thisHiddenNodes = Convert.ToInt32(txtHiddenNodes.Text);
                int thisNFeature = selectedLookBack * NFEATURES;
                network = new ActivationNetwork(
                    new BipolarSigmoidFunction(sigmoidAlphaValue),
                    thisNFeature, thisHiddenNodes, 1);
                networkLoaded = true;
                outS = comboBoxTargets.SelectedIndex.ToString();
                StreamReader sw = new StreamReader(openFileDialog1.FileName);

                string fLine = "";
                fLine = sw.ReadLine(); // read original data file name
                fLine = sw.ReadLine(); // read original target 
                fLine = sw.ReadLine(); // read original # inputs 
                int nLayer = network.Layers.Length;
                int thisNLayer = 0;
                fLine = sw.ReadLine(); // read # of layers
                char[] delimiterChars = { ',' };
                string[] xy = fLine.Split(delimiterChars);
                thisNLayer = Convert.ToInt32(xy[0]);
                if (thisNLayer != nLayer)
                {
                    sw.Close();
                    retVal = "3rd line:  Number of layers doesn't match.  Network = " + nLayer.ToString() + " File = " + thisNLayer.ToString();
                    return retVal;
                }
                MessageBox.Show("Layers = " + nLayer.ToString());
                fLine = sw.ReadLine(); // read # of neurons in each layer (should be nLayer values)
                xy = fLine.Split(delimiterChars);
                if (xy.Length != nLayer)
                {
                    sw.Close();
                    retVal = "4th line:  Number of neurons per layer doesn't match\nNumber of layers: " + nLayer.ToString() + "\nRead: " + xy.Length.ToString() + "\nWeights: " + fLine;
                    return retVal;
                }

                for (int i = 0; i < nLayer; i++)
                {
                    Layer layer = network.Layers[i];
                    int nNeurons = layer.Neurons.Length;
                    for (int j = 0; j < nNeurons; j++)
                    {

                        ActivationNeuron neuron = (ActivationNeuron)layer.Neurons[j];
                        // double threshold = neuron.Threshold;
                        fLine = sw.ReadLine(); // read threshold for ith layer, jth neuron
                        xy = fLine.Split(delimiterChars);
                        double threshold = (double)Convert.ToDouble(xy[0]);
                        neuron.Threshold = threshold;

                        int nWeights = neuron.Weights.Length;
                        fLine = sw.ReadLine(); // read weights for ith layer neurons
                        xy = fLine.Split(delimiterChars);
                        try
                        {
                            for (int k = 0; k < nWeights; k++)
                            {
                                double weight = (double)Convert.ToDouble(xy[k]);
                                neuron.Weights[k] = weight;
                            }
                        }
                        catch (Exception e5)
                        {
                            retVal = "Error reading weights for neurons in layer " + i.ToString() + "\nInput line: " + fLine + "\nError: " + e5.ToString();
                            sw.Close();
                            return retVal;
                        }
                    }

                }
                sw.Close();
                networkLoaded = true;
                comboBoxLookBack.Enabled = true;

            }
            return retVal;
        }


        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show(saveActivationNetworkWeights());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MessageBox.Show(loadActivationNetworkWeights());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MessageBox.Show(saveTestVectors());
        }

        private string saveTestVectors()
        {
            string outS = "";
            string retVal = "Save cancelled";
            saveFileDialog1.DefaultExt = "csv";
            saveFileDialog1.Filter =
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            if (workingDirectory == "")
                saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            else
                saveFileDialog1.InitialDirectory = workingDirectory;
            outS = comboBoxLookBack.SelectedIndex.ToString()+"_"+comboBoxTargets.SelectedIndex.ToString();
            saveFileDialog1.FileName = "TestVectors_" + outS + ".csv"; 
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workingDirectory = Path.GetDirectoryName(saveFileDialog1.FileName);
                retVal = saveFileDialog1.FileName+".csv saved";
                StreamWriter sw = new StreamWriter(saveFileDialog1.FileName+".csv");
                int nTestVectors = listBoxTestVectors.Items.Count;
                for (int i = 0; i < nTestVectors; i++)
                {
                    string str = listBoxTestVectors.Items[i].ToString();
                    string[] xy = str.Split(' ');
                    string csvStr = "";
                    for (int j = 0; j < xy.Length; j++)
                    {
                        csvStr += xy[j] + ",";
                    }
                    str = listBoxTestTargets.Items[i].ToString();
                    xy = str.Split(' ');
                    for (int j = 0; j < xy.Length; j++)
                    {
                        csvStr += xy[j] + ",";
                    }
                    str = listBoxTestPrediction.Items[i].ToString();
                    xy = str.Split(' ');
                    for (int j = 0; j < xy.Length - 1; j++)
                    {
                        csvStr += xy[j] + ",";
                    }
                    csvStr += xy[xy.Length - 1];
                    sw.WriteLine(csvStr);
                }
                sw.Close();
            }
            return retVal;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MessageBox.Show(saveTrainingVectors());
        }

        private string saveTrainingVectors()
        {
            string outS = "";
            string retVal = "Save cancelled";
            saveFileDialog1.DefaultExt = "csv";
            saveFileDialog1.Filter =
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            if (workingDirectory == "")
                saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            else
                saveFileDialog1.InitialDirectory = workingDirectory;
            outS = comboBoxLookBack.SelectedIndex.ToString() + "_" + comboBoxTargets.SelectedIndex.ToString();
            saveFileDialog1.FileName = "TestVectors_" + outS + ".csv";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workingDirectory = Path.GetDirectoryName(saveFileDialog1.FileName);
                retVal = saveFileDialog1.FileName + ".csv saved";
                StreamWriter sw = new StreamWriter(saveFileDialog1.FileName + ".csv");
                int nTrainingVectors = listBoxInputVectors.Items.Count;
                for (int i = 0; i < nTrainingVectors; i++)
                {
                    string str = listBoxInputVectors.Items[i].ToString();
                    string[] xy = str.Split(' ');
                    string csvStr = "";
                    for (int j = 0; j < xy.Length; j++)
                    {
                        csvStr += xy[j] + ",";
                    }
                    str = listBoxTargets.Items[i].ToString();
                    xy = str.Split(' ');
                    for (int j = 0; j < xy.Length; j++)
                    {
                        csvStr += xy[j] + ",";
                    }
                    str = listBoxPrediction.Items[i].ToString();
                    xy = str.Split(' ');
                    for (int j = 0; j < xy.Length - 1; j++)
                    {
                        csvStr += xy[j] + ",";
                    }
                    csvStr += xy[xy.Length - 1];
                    sw.WriteLine(csvStr);
                }
                sw.Close();
            }
            return retVal;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            networkLoaded = true;
            trainNetwork();
        }

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            if (listBoxTrainingSet.Items.Count == 0)
            {
                txtStatus.Text = "Không có dữ liệu huấn luyện";
            }
            else
            {
                btnTrain.Enabled = true;
            }
        }

        private void tabPage4_Enter(object sender, EventArgs e)
        {
            if (listBoxTestSet.Items.Count == 0)
            {
                txtStatus.Text = "Không có dữ liệu kiểm tra";
            }
            else
            {
                btnTest.Enabled = true;
            }

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn thật sự muốn thoát?", "Thoát", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                Application.Exit();
            }

        }

        private void comboBoxOpenIndex_SelectedIndexChanged(object sender, EventArgs e)
        {
            allowedFeatureIndicesSW[0] = comboBoxOpenIndex.SelectedIndex;

        }

        private void comboBoxHighIndex_SelectedIndexChanged(object sender, EventArgs e)
        {
            allowedFeatureIndicesSW[1] = comboBoxHighIndex.SelectedIndex;

        }

        private void comboBoxLowIndex_SelectedIndexChanged(object sender, EventArgs e)
        {
            allowedFeatureIndicesSW[2] = comboBoxLowIndex.SelectedIndex;

        }

        private void comboBoxCloseIndex_SelectedIndexChanged(object sender, EventArgs e)
        {
            allowedFeatureIndicesSW[3] = comboBoxCloseIndex.SelectedIndex;

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                loadData();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Vui lòng nhập dữ liệu");
            }
        }

        private void btnKmean_Click(object sender, EventArgs e)
        {
            

        }

        private void button9_Click(object sender, EventArgs e)
        {
           
        }

        private void comboBoxLookBack_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (txtSoDong.Text != string.Empty )
            {
                int sodong = int.Parse(txtSoDong.Text);

                Random rand = new Random();
                 
                List<int> list = new List<int>();
                for (int i = 0; i < sodong; i++)
                {
                    list.Add(rand.Next(0, listBoxData.Items.Count-1));
                }
                //random
                for (int i = 0; i < list.Count; i++)
                {
                    listBoxData.SetSelected(list[i], true);
                } 
            }
            else
            {
                for (int i = 0; i < listBoxData.Items.Count; i++)
                    listBoxData.SetSelected(i, true);
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {

        }

        private void txtDiem1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void btnDuBao_Click(object sender, EventArgs e)
        {
           if(txtDiem1.Text==string.Empty)
           {
               MessageBox.Show("Bạn cần nhập điểm");
               txtDiem1.Focus();
               return;
           }
           if (txtDiem2.Text == string.Empty)
           {
               MessageBox.Show("Bạn cần nhập điểm");
               txtDiem2.Focus();
               return;
           }
           if (txtDiem3.Text == string.Empty)
           {
               MessageBox.Show("Bạn cần nhập điểm");
               txtDiem3.Focus();
               return;
           }
           if (txtDiem4.Text == string.Empty)
           {
               MessageBox.Show("Bạn cần nhập điểm");
               txtDiem4.Focus();
               return;
           }
           if (txtDiem5.Text == string.Empty)
           {
               MessageBox.Show("Bạn cần nhập điểm");
               txtDiem5.Focus();
               return;
           }
           if (txtDiem6.Text == string.Empty)
           {
               MessageBox.Show("Bạn cần nhập điểm");
               txtDiem6.Focus();
               return;
           }
           
           int nInputVectors =1;
           if (nInputVectors == 0)
           {
               MessageBox.Show("Vui lòng chọn items từ tập huấn luyên sang tập kiểm tra trong tab xử lý dữ liệu");
               return;
           }
            double[][] input = new double[nInputVectors][];
            double[][] output = new double[nInputVectors][];

            int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
            int thisNFeature = selectedLookBack * NFEATURES;

            // create arrays to hold training data and output result
            for (int i = 0; i < nInputVectors; i++)
            {
                input[i] = new double[thisNFeature];
                output[i] = new double[1];
            }
         //   data = new double[NFEATURES, MAXSAMPLES];
            // populate the input vector with items from the test list taken LookBack at a time
            string str = null;
            char[] delimiterChars = { ' ' };
            int nVectors = 0;
            int n = 0;
            for (int i = 0; i < nInputVectors; i++)//for (int i = 0; i < nInputVectors - thisNFeature - 1; i++)
            {
                str = "1 2001110022 "+txtDiem1.Text+" "+txtDiem2.Text+" "+txtDiem3.Text+" "+txtDiem4.Text+" "+txtDiem5.Text+" "+txtDiem6.Text+" 10.0";//listBoxTestSet.Items[i].ToString();
                string[] xy = str.Split(delimiterChars);
                int thisInd = Convert.ToInt32(xy[0]); // index into data array
                n = 0;
                for (int j = 0; j < selectedLookBack; j++)
                {
                    for (int k = 0; k < NFEATURES-1; k++)
                    {
                        input[i][n] = double.Parse(xy[k+2]);
                        n++;
                    }
                }
                //  MessageBox.Show(comboBoxTargets.SelectedItem.ToString());
                int outInd = Convert.ToInt32(comboBoxTargets.SelectedIndex);
             // output[i][0] = data[outInd, thisInd + selectedLookBack - 1];
               // outputTestIndex[i] = thisInd + selectedLookBack;
                nVectors++;
            }
            scaleTestIO_SV(nVectors, thisNFeature, input, output);
           
            double[] networkInput = new double[thisNFeature];
       
            for (int i = 0; i < nVectors; i++)
            {
                // put values from current window as network's input
                for (int j = 0; j < thisNFeature; j++)
                {
                    networkInput[j] = input[i][j];
                }
               // double learningErrorTmp = 0D;

                double[] solN = network.Compute(networkInput); // + 0.85) / factor + yMin;
            
              
            
                double tmin = minTestValues[i];
                double diff = maxTestValues[i] - tmin;
                double iv = (solN[0] + .9D) / 1.8D * diff + tmin;
               
                double aa = iv - errorToancuc;
                lblDuBaoSV.Text = iv.ToString("#.##");
                lblXepLoai.Text = XepLoai(iv);
                if (lblXepLoai.Text== "Giỏi")
                {
                    lblDuBaoSV.ForeColor = Color.Lime;
                    lblXepLoai.ForeColor = Color.Lime;
                }
                else if(lblXepLoai.Text== "Khá")
                {
                    lblDuBaoSV.ForeColor = Color.Orange;
                    lblXepLoai.ForeColor = Color.Orange;
                }else if(lblXepLoai.Text== "Trung bình")
                {
                    lblDuBaoSV.ForeColor = Color.Red;
                    lblXepLoai.ForeColor = Color.Red;
                }

            }
        }

        public string XepLoai(double pInput)
        {
            if (pInput > 8.4)
                return "Giỏi";
            if (pInput >= 7 && pInput <= 8.4)
                return "Khá";
            if (pInput < 4)
                return "Yếu";
            if (pInput >= 4 && pInput < 7)
                return "Trung bình";
            return "";
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamReader reader = null;
                // read maximum MAXSAMPLES points
                rawInputFileName = openFileDialog1.FileName;
                try
                {
                    // open selected file
                    reader = File.OpenText(rawInputFileName);
                    string str = null;
                    char[] delimiterChars = { ',' };


                    if ((str = reader.ReadLine()) != null) // feature names
                    {
                        string[] xy = str.Split(delimiterChars);
                        int nT = xy.Length;
                       
                    }
                    txtStatus.Text = "Modify columns to identify OHLC if necessary and then click \"Load Data\" button";


                }
                catch (Exception e7)
                {
                    MessageBox.Show("Failed reading the file\n" + e7.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    // close file
                    if (reader != null)
                        reader.Close();
                }
                loadDataGV();
            }
        }
        private void loadDataGV()
        {
            data_GV = new double[NFEATURES, MAXSAMPLES]; // kích thướt ma trận data
            StreamReader reader = null;
            
            nData = 0;

            lstGVInput.Items.Clear();
            try
            {
                // open selected file
                reader = File.OpenText(rawInputFileName);
                string str = null;
                char[] delimiterChars = { ',' };

                if ((str = reader.ReadLine()) != null) // feature names
                {
                    string[] xy = str.Split(delimiterChars);
                    int nT = xy.Length;
                     
                }
                else
                {
                    txtStatus.Text = "Couldn't read first line (feature/target names)";
                    return;
                }
                nData = 1;
                while ((str = reader.ReadLine()) != null) // read values
                {

                    string[] xy = str.Split(delimiterChars);
                    string dt = xy[1];
                    for (int j = 0; j < nFeatures; j++) // feature values
                        data_GV[j, nData] = double.Parse(xy[allowedFeatureIndicesSW[j]]);
                    string lbSTR = nData.ToString() + " " + dt + " " + data_GV[0, nData] + " " + data_GV[1, nData] + " " + data_GV[2, nData] + " " + data_GV[3, nData]
                        + " " + data_GV[4, nData] + " " + data_GV[5, nData] + " " + data_GV[6, nData];
                    lstGVInput.Items.Add(lbSTR);
                    nData++;

                }
                txtStatus.Text = "Đọc " + nData.ToString() + " dữ liệu dự báo từ " + Path.GetFileName(rawInputFileName);

            }
            catch (Exception e7)
            {
                MessageBox.Show("Lỗi đọc file\n" + e7.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                // close file
                if (reader != null)
                    reader.Close();
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {

            int nInputVectors = lstGVInput.Items.Count;
            if (nInputVectors == 0)
            {
                MessageBox.Show("Vui lòng chọn items từ tập huấn luyên sang tập kiểm tra trong tab xử lý dữ liệu");
                return;
            }
            double[][] input = new double[nInputVectors][];
            double[][] output = new double[nInputVectors][];

            int selectedLookBack = Convert.ToInt32(comboBoxLookBack.SelectedItem);
            int thisNFeature = selectedLookBack * NFEATURES;

            // create arrays to hold training data and output result
            for (int i = 0; i < nInputVectors; i++)
            {
                input[i] = new double[thisNFeature];
                output[i] = new double[1];
            }

            // populate the input vector with items from the test list taken LookBack at a time
            string str = null;
            char[] delimiterChars = { ' ' };
            int nVectors = 0;
            int n = 0;
            for (int i = 0; i < nInputVectors; i++)//for (int i = 0; i < nInputVectors - thisNFeature - 1; i++)
            {
                str = lstGVInput.Items[i].ToString();
                string[] xy = str.Split(delimiterChars);
                int thisInd = Convert.ToInt32(xy[0]); // index into data array
                n = 0;
                for (int j = 0; j < selectedLookBack; j++)
                {
                    for (int k = 0; k < NFEATURES; k++)
                    {
                         input[i][n] = double.Parse( xy[k+2]);
                        n++;
                    }
                }
         
                nVectors++;
            }
            scaleTestIO_GV(nVectors, thisNFeature, input, output);
       
            double[] networkInput = new double[thisNFeature];
          
            lstKetquaGV.Items.Clear();
          
            for (int i = 0; i < nVectors; i++)
            {
                // put values from current window as network's input
                for (int j = 0; j < thisNFeature; j++)
                {
                    networkInput[j] = input[i][j];
                }

                double[] solN = network.Compute(networkInput); // + 0.85) / factor + yMin;
             
                double tmin = minTestValues[i];
                double diff = maxTestValues[i] - tmin;
                double iv = (solN[0] + .9D) / 1.8D * diff + tmin;
                /// double iv = solN[0];
            
                lstKetquaGV.Items.Add((i+1).ToString()+" " + iv.ToString("#.##"));
                 
               

            }

        }

        private void button15_Click(object sender, EventArgs e)
        {
            listBox4.Items.Clear();
            int gioi = 0;
            int kha = 0;
            int tb = 0;
            int yeu = 0;
            for (int i = 0; i < lstKetquaGV.Items.Count; i++)
            {
                string[] xy = lstKetquaGV.Items[i].ToString().Split(' ');
                double pInput = double.Parse(xy[1]);
                if (pInput > 8.4)
                    gioi++;
                if (pInput >= 7 && pInput <= 8.4)
                    kha++;
                if (pInput < 4)
                    yeu++;
                if (pInput >= 4 && pInput < 7)
                    tb++;
            }
            listBox4.Items.Add("Giỏi:"+gioi);
            listBox4.Items.Add("Khá:" + kha);
            listBox4.Items.Add("TB:" + tb);
            listBox4.Items.Add("Yếu:" + yeu);
        }

        private void txtSoDong_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(!char.IsDigit(e.KeyChar)&&!char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }
            //xuat Excel

        private void btn_XuatBC_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xls)|*.xls",
                    FilterIndex = 2,
                    RestoreDirectory = true
                };
                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path = saveFileDialog.FileName;
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }
                    int count = 0;
                    Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlWorkbook = xlApp.Workbooks.Add(misValue);

                    xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                    xlWorksheet.Rows["1"].Font.Color = Color.Red;
                    xlWorksheet.Rows["3"].Font.Color = Color.Green;
                    xlWorksheet.Cells[1, 3] = "DỮ LIỆU DỰ BÁO";
                    xlWorksheet.Cells[1, 9] = "KẾT QUẢ DỰ BÁO";
                    xlWorksheet.Cells[1, 12] = "KẾT QUẢ THỐNG KÊ";

                    xlWorksheet.Cells[3, 3] = "Stt| MSSV |M1|M2|M3|M4|M5|M6|M7";
                    xlWorksheet.Cells[3, 9] = "Stt | MDB";

                    foreach (var item in lstGVInput.Items)
                    {
                        xlWorksheet.Cells[5 + count, 3] = item;
                        count++;
                    }
                    count = 0;
                    foreach (var item in lstKetquaGV.Items)
                    {
                        xlWorksheet.Cells[5 + count, 9] = item;
                        count++;
                    }
                    count = 0;
                    foreach (var item in listBox4.Items)
                    {
                        xlWorksheet.Cells[5 + count, 12] = item;
                        count++;
                    }

                    //End
                    Microsoft.Office.Interop.Excel.Range chartRange = xlWorksheet.get_Range("C5", "F474");
                    Microsoft.Office.Interop.Excel.Range chartRange_2 = xlWorksheet.get_Range("I5", "I474");
                    Microsoft.Office.Interop.Excel.Range chartRange_3 = xlWorksheet.get_Range("L5", "L8");
                    Microsoft.Office.Interop.Excel.Range chartRange_4 = xlWorksheet.get_Range("C3", "F3");
                    Microsoft.Office.Interop.Excel.Range chartRange_5 = xlWorksheet.get_Range("I3");

                    chartRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    chartRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    chartRange_2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    chartRange_2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    chartRange_3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    chartRange_3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    chartRange_4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    chartRange_4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    chartRange_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    chartRange_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    xlWorkbook.SaveAs(path,
                                    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                                    misValue, misValue, misValue, misValue,
                                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                                    misValue, misValue, misValue, misValue, misValue);

                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorksheet);
                    Marshal.ReleaseComObject(xlWorkbook);
                    Marshal.ReleaseComObject(xlApp);
                    MessageBox.Show("Xuất File Excel thành công !");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            WindowState= FormWindowState.Minimized;
        }
    }
}
