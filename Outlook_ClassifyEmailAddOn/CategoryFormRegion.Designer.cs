namespace Outlook_ClassifyEmailAddOn
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class CategoryFormRegion : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        private System.Windows.Forms.Button button_trainModel;

        private System.Windows.Forms.GroupBox groupBox_classifications;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        public System.Windows.Forms.Button button_read;
        private System.Windows.Forms.Button button_ignore;
        private System.Windows.Forms.Button button_delete;
        private System.Windows.Forms.Button button_followUp;
        private string prediction;


        public CategoryFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            :
            base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button_trainModel = new System.Windows.Forms.Button();
            this.groupBox_classifications = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_read = new System.Windows.Forms.Button();
            this.button_ignore = new System.Windows.Forms.Button();
            this.button_delete = new System.Windows.Forms.Button();
            this.button_followUp = new System.Windows.Forms.Button();
            this.groupBox_classifications.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();

            // button_trainModel
            // 
            this.button_trainModel.Location = new System.Drawing.Point(1035, 12);
            this.button_trainModel.Name = "button_trainModel";
            this.button_trainModel.Size = new System.Drawing.Size(75, 23);
            this.button_trainModel.TabIndex = 2;
            this.button_trainModel.Text = "Train Model";
            this.button_trainModel.UseVisualStyleBackColor = true;
            this.button_trainModel.Click += new System.EventHandler(this.Button_trainModel_Click);
            // 
            // groupBox_classifications
            // 
            this.groupBox_classifications.Controls.Add(this.tableLayoutPanel1);
            this.groupBox_classifications.Location = new System.Drawing.Point(24, 37);
            this.groupBox_classifications.Name = "groupBox_classifications";
            this.groupBox_classifications.Size = new System.Drawing.Size(1101, 115);
            this.groupBox_classifications.TabIndex = 3;
            this.groupBox_classifications.TabStop = false;
            this.groupBox_classifications.Text = "Select a folder for Classifications";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 224F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 209F));
            this.tableLayoutPanel1.Controls.Add(this.button_read, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.button_ignore, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button_delete, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.button_followUp, 1, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(17, 30);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1078, 79);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // button_read
            // 
            this.button_read.Location = new System.Drawing.Point(3, 3);
            this.button_read.Name = "button_read";
            this.button_read.Size = new System.Drawing.Size(75, 23);
            this.button_read.TabIndex = 7;
            this.button_read.Text = "Read";
            this.button_read.UseVisualStyleBackColor = false;
            this.button_read.Click += new System.EventHandler(this.button_read_Click);
            // 
            // button_ignore
            // 
            this.button_ignore.Location = new System.Drawing.Point(3, 42);
            this.button_ignore.Name = "button_ignore";
            this.button_ignore.Size = new System.Drawing.Size(75, 23);
            this.button_ignore.TabIndex = 6;
            this.button_ignore.Text = "Ignore";
            this.button_ignore.UseVisualStyleBackColor = true;
            this.button_ignore.Click += new System.EventHandler(this.button_ignore_Click);
            // 
            // button_delete
            // 
            this.button_delete.Location = new System.Drawing.Point(430, 3);
            this.button_delete.Name = "button_delete";
            this.button_delete.Size = new System.Drawing.Size(75, 23);
            this.button_delete.TabIndex = 5;
            this.button_delete.Text = "Delete";
            this.button_delete.UseVisualStyleBackColor = true;
            this.button_delete.Click += new System.EventHandler(this.button_delete_Click);
            // 
            // button_followUp
            // 
            this.button_followUp.Location = new System.Drawing.Point(430, 42);
            this.button_followUp.Name = "button_followUp";
            this.button_followUp.Size = new System.Drawing.Size(75, 23);
            this.button_followUp.TabIndex = 4;
            this.button_followUp.Text = "Follow Up";
            this.button_followUp.UseVisualStyleBackColor = true;
            this.button_followUp.Click += new System.EventHandler(this.button_followUp_Click);

            // 
            // preicition
            // 
            this.prediction = Globals.ThisAddIn.classifyEmail();
            switch (this.prediction)
            {
                case "Read":
                    this.button_read.BackColor = System.Drawing.SystemColors.Highlight;
                    break;
                case "Ignore":
                    this.button_ignore.BackColor = System.Drawing.SystemColors.Highlight;
                    break;
                case "Delete":
                    this.button_delete.BackColor = System.Drawing.SystemColors.Highlight;
                    break;
                case "FollowUp":
                    this.button_followUp.BackColor = System.Drawing.SystemColors.Highlight;
                    break;
                default:
                    break;
            }
            // 
            // ML_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.groupBox_classifications);
            this.Controls.Add(this.button_trainModel);
            this.Name = "ML_Form";
            this.Size = new System.Drawing.Size(1128, 162);
            this.FormRegionShowing += new System.EventHandler(this.FormRegion1_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.FormRegion1_FormRegionClosed);
            this.groupBox_classifications.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        #region Form Region Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            manifest.FormRegionName = "FormRegion1";
            manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining;

        }

        #endregion



        public partial class FormRegion1Factory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public FormRegion1Factory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                CategoryFormRegion.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.FormRegion1Factory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                CategoryFormRegion form = new CategoryFormRegion(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal CategoryFormRegion FormRegion1
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(CategoryFormRegion))
                        return (CategoryFormRegion)item;
                }
                return null;
            }
        }
    }
}
