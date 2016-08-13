using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;
using Microsoft.VisualBasic.Devices;
using Microsoft.Win32;
using System.Windows;
using System.Management;
using System.IO;
namespace MyPcPassport
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            refreshToolStripMenuItem.Enabled = false;
            new Thread(() =>
            {
                //Create basic variables 
                Computer cpu = new Computer();
                TreeNode Sys = new TreeNode(Environment.UserName + "@" + Environment.MachineName);

                //SOFTWARE NODE 
                TreeNode OS = new TreeNode("OS");
                OS.Nodes.Add("Operating System");
                OS.Nodes.Add("Windows: " + cpu.Info.OSFullName);
                OS.Nodes.Add("Is64Bit: " + Environment.Is64BitOperatingSystem);
                OS.Nodes.Add("Service Pack: " + Environment.OSVersion.ServicePack);
                OS.Nodes.Add("Windows Version: " + cpu.Info.OSVersion);
                OS.Nodes.Add("Windows Platform: " + cpu.Info.OSPlatform);
                UpdateProgress(10);
                Thread.Sleep(150);

                //CPU NODE 
                TreeNode CPU = new TreeNode("CPU");
                CPU.Nodes.Add("Central Processing Unit");
                CPU.Nodes.Add("Name: " + Processor.GetProcessorType());
                CPU.Nodes.Add("Cores: " + Environment.ProcessorCount);
                CPU.Nodes.Add("Identifier: " + Processor.GetVendor());
                CPU.Nodes.Add("Vendor: " + Processor.GetIdentifier());
                UpdateProgress(20);
                Thread.Sleep(150);

                //GPU NODE
                TreeNode GPU = new TreeNode("GPU");
                GPU.Nodes.Add("Graphics Processing Unit");
                GPU.Nodes.Add("Name: " + GraphicsProcessor.GetGraphicsCard());
                UpdateProgress(30);
                Thread.Sleep(150);

                //RAM NODE 
                TreeNode RAM = new TreeNode("RAM");
                RAM.Nodes.Add("Random Access Memory");
                RAM.Nodes.Add(string.Format("Used RAM: {0:0.00}MB", ((cpu.Info.TotalPhysicalMemory - cpu.Info.AvailablePhysicalMemory) / 1024) / 1024));
                RAM.Nodes.Add(string.Format("Available RAM: {0:0.00}MB", (cpu.Info.AvailablePhysicalMemory / 1024) / 1024));
                RAM.Nodes.Add(string.Format("Total RAM: {0:0.00}MB", (cpu.Info.TotalPhysicalMemory / 1024) / 1024));
                UpdateProgress(40);
                Thread.Sleep(150);

                //HDD NODE 
                TreeNode HDD = new TreeNode("HDD");
                HDD.Nodes.Add("Hard Disk Drives & Other Internal Storage");
                int count = 0;
                foreach(DriveInfo di in DriveInfo.GetDrives())
                {
                    if(di.DriveType == DriveType.Fixed)
                    {
                        count++;
                        TreeNode slot = new TreeNode("SLOT " + count);
                        slot.Nodes.Add("Label: " + di.VolumeLabel);
                        slot.Nodes.Add("Name: " + di.Name);
                        slot.Nodes.Add("Root: " + di.RootDirectory);
                        slot.Nodes.Add("Format: " + di.DriveFormat);
                        slot.Nodes.Add("Available Free Space: " + ((di.AvailableFreeSpace / 1024) / 1024) / 1024 + "GB");
                        slot.Nodes.Add("Total Free Space: " + ((di.TotalFreeSpace / 1024) / 1024) / 1024 + "GB");
                        slot.Nodes.Add("Total Space: " + ((di.TotalSize / 1024) / 1024) / 1024 + "GB");
                        HDD.Nodes.Add(slot);
                    }

                }
                UpdateProgress(75);
                Thread.Sleep(500);
                

                //Add Nodes to root node 
                Sys.Nodes.Add(OS);
                Sys.Nodes.Add(CPU);
                Sys.Nodes.Add(GPU);
                Sys.Nodes.Add(RAM);
                Sys.Nodes.Add(HDD);

                //Complete the system info grabing 
                Invoke((MethodInvoker)delegate { this.treeView1.Nodes.Add(Sys); refreshToolStripMenuItem.Enabled = true; });
                UpdateProgress(100);
            }).Start();
        }

        public void UpdateProgress(int progress)
        {
            if(InvokeRequired)
                Invoke((MethodInvoker)delegate { UpdateProgress(progress); });
            else 
                this.pbProgress.Value = progress;
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            refreshToolStripMenuItem.Enabled = false;
            treeView1.Nodes.Clear();
            new Thread(() =>
            {
                //Create basic variables 
                Computer cpu = new Computer();
                TreeNode Sys = new TreeNode(Environment.UserName + "@" + Environment.MachineName);

                //SOFTWARE NODE 
                TreeNode OS = new TreeNode("OS");
                OS.Nodes.Add("Operating System");
                OS.Nodes.Add("Windows: " + cpu.Info.OSFullName);
                OS.Nodes.Add("Is64Bit: " + Environment.Is64BitOperatingSystem);
                OS.Nodes.Add("Service Pack: " + Environment.OSVersion.ServicePack);
                OS.Nodes.Add("Windows Version: " + cpu.Info.OSVersion);
                OS.Nodes.Add("Windows Platform: " + cpu.Info.OSPlatform);
                UpdateProgress(10);
                Thread.Sleep(150);

                //CPU NODE 
                TreeNode CPU = new TreeNode("CPU");
                CPU.Nodes.Add("Central Processing Unit");
                CPU.Nodes.Add("Name: " + Processor.GetProcessorType());
                CPU.Nodes.Add("Cores: " + Environment.ProcessorCount);
                CPU.Nodes.Add("Identifier: " + Processor.GetVendor());
                CPU.Nodes.Add("Vendor: " + Processor.GetIdentifier());
                UpdateProgress(20);
                Thread.Sleep(150);

                //GPU NODE
                TreeNode GPU = new TreeNode("GPU");
                GPU.Nodes.Add("Graphics Processing Unit");
                GPU.Nodes.Add("Name: " + GraphicsProcessor.GetGraphicsCard());
                UpdateProgress(30);
                Thread.Sleep(150);

                //RAM NODE 
                TreeNode RAM = new TreeNode("RAM");
                RAM.Nodes.Add("Random Access Memory");
                RAM.Nodes.Add(string.Format("Used RAM: {0:0.00}MB", ((cpu.Info.TotalPhysicalMemory - cpu.Info.AvailablePhysicalMemory) / 1024) / 1024));
                RAM.Nodes.Add(string.Format("Available RAM: {0:0.00}MB", (cpu.Info.AvailablePhysicalMemory / 1024) / 1024));
                RAM.Nodes.Add(string.Format("Total RAM: {0:0.00}MB", (cpu.Info.TotalPhysicalMemory / 1024) / 1024));
                UpdateProgress(40);
                Thread.Sleep(150);

                //HDD NODE 
                TreeNode HDD = new TreeNode("HDD");
                HDD.Nodes.Add("Hard Disk Drives & Other Internal Storage");
                int count = 0;
                foreach (DriveInfo di in DriveInfo.GetDrives())
                {
                    if (di.DriveType == DriveType.Fixed)
                    {
                        count++;
                        TreeNode slot = new TreeNode("SLOT " + count);
                        slot.Nodes.Add("Label: " + di.VolumeLabel);
                        slot.Nodes.Add("Name: " + di.Name);
                        slot.Nodes.Add("Root: " + di.RootDirectory);
                        slot.Nodes.Add("Format: " + di.DriveFormat);
                        slot.Nodes.Add("Available Free Space: " + ((di.AvailableFreeSpace / 1024) / 1024) / 1024 + "GB");
                        slot.Nodes.Add("Total Free Space: " + ((di.TotalFreeSpace / 1024) / 1024) / 1024 + "GB");
                        slot.Nodes.Add("Total Space: " + ((di.TotalSize / 1024) / 1024) / 1024 + "GB");
                        HDD.Nodes.Add(slot);
                    }

                }
                UpdateProgress(75);
                Thread.Sleep(500);


                //Add Nodes to root node 
                Sys.Nodes.Add(OS);
                Sys.Nodes.Add(CPU);
                Sys.Nodes.Add(GPU);
                Sys.Nodes.Add(RAM);
                Sys.Nodes.Add(HDD);

                //Complete the system info grabing 
                Invoke((MethodInvoker)delegate { this.treeView1.Nodes.Add(Sys); refreshToolStripMenuItem.Enabled = true; });
                UpdateProgress(100);
            }).Start();
            
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = ".txt|*.txt";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (System.IO.StreamWriter sw = new StreamWriter(sfd.FileName))
                {
                    sw.WriteLine("MyPcPassport 1.0");
                    sw.WriteLine("(C) JordanHook.com 2014");
                    sw.WriteLine("---");
                    sw.WriteLine(treeView1.Nodes[0].Text); //Write pc name 

                    //Loop through all hardware components
                    foreach(TreeNode node in treeView1.Nodes[0].Nodes)
                    {
                        //list the component 
                        sw.WriteLine("" + node.Text);

                        //List information about said component 
                        foreach(TreeNode info in node.Nodes)
                        {
                            sw.WriteLine("\t" + info.Text);

                            //some components are listed as a secondary node, so lets loop once more... 
                            foreach(TreeNode n in info.Nodes)
                            {
                                sw.WriteLine("\t\t" + n.Text);
                            }
                        }
                    }
                }
            }
            sfd.Dispose();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(-2);
        }

        private void fAQToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void websiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //launch website 
            System.Diagnostics.Process.Start("http://jordanhook.com");
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //short message about the software 
            MessageBox.Show("MyPcPassport was developed by Jordan Hook of JordanHook.com in C#. The purpose of the application is to be a light weight, system information and dianostics tool in order to help people identify their hardware and system specifications. From time to time, the project will be updated so be sure to check out our website for updates and other cool software!");
        }


        bool expand = true;
        private void expandAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //use the class variable to detect if it's already expanded and do the opposite 
            if (expand)
                treeView1.ExpandAll();
            else
                treeView1.CollapseAll();

            //Switch the class variable to match what action needs to be done for next time
            expand = !expand;
        }
    }
    
    class GraphicsProcessor
    {
        public static string GetGraphicsCard()
        {
            try
            {
                //Query the system for the graphics card information 
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DisplayConfiguration");
                foreach (ManagementObject mo in searcher.Get())
                {
                    //if a result is found... check the properties 
                    foreach (PropertyData property in mo.Properties)
                    {
                        //if description is found... 
                        if (property.Name == "Description")
                        {
                            //return the description of the GPU 
                            return property.Value.ToString();
                        }
                    }
                }

                //other wise return error 
                return "Error!";
            }
            catch { return "Error!"; }
        }

    }

    class Processor
    {
        public static string GetVendor()
        {
            try
            {
                //Access registry to locate information about the proecessor 
                RegistryKey RegKey = Registry.LocalMachine;
                RegKey = RegKey.OpenSubKey("HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0");
                return (string)RegKey.GetValue("Identifier");
            }
            catch { return "Error!"; }
        }
        public static string GetIdentifier()
        {
            try
            {
                //Access registry to locate information about the proecessor 
                RegistryKey RegKey = Registry.LocalMachine;
                RegKey = RegKey.OpenSubKey("HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0");
                return (string)RegKey.GetValue("VendorIdentifier");
            }
            catch { return "Error!"; }
        }

        public static string GetProcessorType()
        {
            try
            {
                //Access registry to locate information about the proecessor 
                RegistryKey RegKey = Registry.LocalMachine;
                RegKey = RegKey.OpenSubKey("HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0");
                return (string)RegKey.GetValue("ProcessorNameString");
            }
            catch { return "Error!"; }
        }
    }
}
