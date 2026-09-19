using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: AssemblyTitle("Outlook Part Sender")]
[assembly: AssemblyDescription("Outlook Classic Portable Outlook helper for sending attachments in numbered parts")]
[assembly: AssemblyCompany("Local")]
[assembly: AssemblyProduct("Outlook Part Sender")]
[assembly: AssemblyVersion("3.1.0.0")]
[assembly: AssemblyFileVersion("3.1.0.0")]
[assembly: ComVisible(false)]

namespace OutlookPartSender
{
    [ComVisible(true)]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    [ComVisible(true)]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    [ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    [ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    [ComVisible(true)]
    [Guid("A93D78E5-9A64-4D49-9B9F-85B0EA75C0A1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IOutlookPartSenderCallbacks
    {
        [DispId(1)]
        void OnSendParts([MarshalAs(UnmanagedType.IDispatch)] object control);
    }

    [ComVisible(true)]
    [Guid("74A51DAB-B18E-4E09-B9D3-72DAEF92C1DE")]
    [ProgId("OutlookPartSender.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IOutlookPartSenderCallbacks))]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility, IOutlookPartSenderCallbacks
    {
        private object _outlookApplication;
        private PartSenderForm _form;
        private volatile bool _hostShuttingDown;

        public Connect()
        {
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _outlookApplication = Application;
                _hostShuttingDown = false;
                Log("OnConnection: " + ConnectMode.ToString());
            }
            catch (Exception ex)
            {
                Log("OnConnection ERROR: " + ex.ToString());
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                _hostShuttingDown = true;
                CloseFormForHostShutdown();
                _outlookApplication = null;
                Log("OnDisconnection: " + RemoveMode.ToString());
            }
            catch (Exception ex)
            {
                Log("OnDisconnection ERROR: " + ex.ToString());
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                _hostShuttingDown = true;
                CloseFormForHostShutdown();
                Log("OnBeginShutdown");
            }
            catch (Exception ex)
            {
                Log("OnBeginShutdown ERROR: " + ex.ToString());
            }
        }

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                if (!String.Equals(RibbonID, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
                    return String.Empty;

                return
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"tabOutlookPartSender\" label=\"Part Sender\">" +
                    "        <group id=\"grpOutlookPartSender\" label=\"Documentos em Partes\">" +
                    "          <button id=\"btnOutlookPartSender\" label=\"Enviar em Partes\" size=\"large\" " +
                    "                  screentip=\"Enviar anexos em vários e-mails numerados\" " +
                    "                  supertip=\"Monte quantas partes precisar e associe um ou vários anexos a cada e-mail.\" " +
                    "                  onAction=\"OnSendParts\"/>" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";
            }
            catch (Exception ex)
            {
                Log("GetCustomUI ERROR: " + ex.ToString());
                return String.Empty;
            }
        }

        public void OnSendParts(object control)
        {
            try
            {
                if (_hostShuttingDown || _outlookApplication == null)
                    return;

                if (_form != null && !_form.IsDisposed)
                {
                    if (!_form.Visible) _form.Show();
                    if (_form.WindowState == FormWindowState.Minimized) _form.WindowState = FormWindowState.Normal;
                    _form.BringToFront();
                    _form.Activate();
                    return;
                }

                _form = new PartSenderForm(_outlookApplication, IsHostShuttingDown);
                _form.FormClosed += delegate(object sender, FormClosedEventArgs e)
                {
                    _form = null;
                };
                _form.Show();
                _form.BringToFront();
                _form.Activate();
            }
            catch (Exception ex)
            {
                Log("OnSendParts ERROR: " + ex.ToString());
                MessageBox.Show(
                    "Não foi possível abrir o Outlook Part Sender.\r\n\r\n" + ex.Message,
                    "Outlook Part Sender",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private bool IsHostShuttingDown()
        {
            return _hostShuttingDown;
        }

        private void CloseFormForHostShutdown()
        {
            if (_form == null || _form.IsDisposed) return;

            try
            {
                _form.NotifyHostShutdown();
            }
            catch
            {
            }
        }

        internal static void Log(string message)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookPartSender");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "addin.log");
                File.AppendAllText(
                    path,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch
            {
            }
        }
    }

    internal sealed class AccountInfo
    {
        public object ComObject;
        public string Display;
        public string Key;

        public override string ToString()
        {
            return Display ?? "";
        }
    }

    internal sealed class PartModel
    {
        public readonly List<string> Attachments = new List<string>();
    }

    internal sealed class PartSnapshot
    {
        public string[] Attachments;
    }

    internal sealed class BatchSnapshot
    {
        public AccountInfo Account;
        public string To;
        public string Cc;
        public string Bcc;
        public string Subject;
        public string Body;
        public bool IncludeSignature;
        public bool PartInBody;
        public bool SendNow;
        public bool SizeWarning;
        public int SizeWarningMb;
        public List<PartSnapshot> Parts;
    }

    internal sealed class SetupConfig
    {
        public string Version { get; set; }
        public string AccountKey { get; set; }
        public string To { get; set; }
        public string CC { get; set; }
        public string BCC { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public bool IncludeSignature { get; set; }
        public bool PartInBody { get; set; }
        public bool SizeWarning { get; set; }
        public int SizeWarningMB { get; set; }
        public List<SetupPart> Parts { get; set; }
    }

    internal sealed class SetupPart
    {
        public List<string> Attachments { get; set; }
    }

    internal sealed class ReviewForm : Form
    {
        public ReviewForm(string text)
        {
            Text = "Revisão — Parte → Anexos";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(900, 650);
            MinimumSize = new Size(700, 500);
            AutoScaleMode = AutoScaleMode.Dpi;

            TextBox box = new TextBox();
            box.Multiline = true;
            box.ReadOnly = true;
            box.ScrollBars = ScrollBars.Both;
            box.WordWrap = false;
            box.Dock = DockStyle.Fill;
            box.Font = new Font("Consolas", 9.0f);
            box.Text = text;

            Button close = new Button();
            close.Text = "Fechar";
            close.Dock = DockStyle.Bottom;
            close.Height = 38;
            close.DialogResult = DialogResult.OK;

            Controls.Add(box);
            Controls.Add(close);
            AcceptButton = close;
            CancelButton = close;
        }
    }

    internal sealed class PartSenderForm : Form
    {
        private readonly object _outlook;
        private readonly Func<bool> _hostShuttingDown;
        private readonly List<PartModel> _parts = new List<PartModel>();
        private readonly List<AccountInfo> _accounts = new List<AccountInfo>();

        private bool _busy;
        private bool _dirty;
        private bool _hostShutdown;
        private bool _loading;

        private ComboBox cmbAccount;
        private TextBox txtTo;
        private TextBox txtCC;
        private TextBox txtBCC;
        private TextBox txtSubject;
        private RichTextBox txtBody;
        private ListView lvParts;
        private ListView lvAttachments;
        private Label lblSelectedPart;
        private Label lblPreview;
        private Label lblCount;
        private CheckBox chkSignature;
        private CheckBox chkPartInBody;
        private CheckBox chkSizeWarning;
        private NumericUpDown numMaxMB;
        private Label lblSizeWarning;
        private RadioButton rbDraft;
        private RadioButton rbSend;
        private Label lblStatus;

        private Button btnReload;
        private Button btnFilesToParts;
        private Button btnNewPart;
        private Button btnDeletePart;
        private Button btnPartUp;
        private Button btnPartDown;
        private Button btnAddAttachment;
        private Button btnRemoveAttachment;
        private Button btnSortAttachments;
        private Button btnSave;
        private Button btnLoad;
        private Button btnReview;
        private Button btnGenerate;

        public PartSenderForm(object outlookApplication, Func<bool> hostShuttingDown)
        {
            _outlook = outlookApplication;
            _hostShuttingDown = hostShuttingDown;

            BuildUi();
            WireEvents();
            LoadAccounts(null);
            RefreshPartsView(-1);

            _dirty = false;
        }

        public void NotifyHostShutdown()
        {
            _hostShutdown = true;

            try
            {
                Enabled = false;
                lblStatus.Text = "Outlook está encerrando...";
            }
            catch
            {
            }

            if (!_busy)
            {
                try
                {
                    Close();
                }
                catch
                {
                }
            }
        }

        private void BuildUi()
        {
            Text = "Outlook Part Sender v3.1.0 PORTABLE";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1180, 930);
            MinimumSize = new Size(1000, 760);
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScroll = true;
            Font = new Font("Segoe UI", 9.0f);

            Label title = NewLabel("Envio de documentação em partes", 24, 16, 800, 34);
            title.Font = new Font("Segoe UI", 16.0f, FontStyle.Bold);

            Label subtitle = NewLabel(
                "Quantidade livre de partes. Cada Parte pode receber um ou vários anexos específicos.",
                27, 52, 950, 24);

            NewLabel("Conta de envio:", 28, 92, 110, 22);
            cmbAccount = new ComboBox();
            cmbAccount.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbAccount.Location = new Point(145, 88);
            cmbAccount.Size = new Size(760, 28);
            Controls.Add(cmbAccount);

            btnReload = NewButton("Atualizar contas", 920, 87, 150, 31);

            NewLabel("Para:", 28, 133, 110, 22);
            txtTo = NewTextBox(145, 129, 925);

            NewLabel("CC:", 28, 169, 110, 22);
            txtCC = NewTextBox(145, 165, 925);

            NewLabel("BCC (opcional):", 28, 205, 110, 22);
            txtBCC = NewTextBox(145, 201, 925);

            NewLabel("Assunto base:", 28, 241, 110, 22);
            txtSubject = NewTextBox(145, 237, 925);

            NewLabel("Mensagem:", 28, 279, 110, 22);
            txtBody = new RichTextBox();
            txtBody.Location = new Point(145, 275);
            txtBody.Size = new Size(925, 100);
            Controls.Add(txtBody);

            Label hint = NewLabel(
                "Destinatários múltiplos: separe por ponto e vírgula (;). A mensagem é repetida em todas as partes.",
                145, 379, 900, 22);
            hint.ForeColor = SystemColors.GrayText;

            GroupBox grpParts = new GroupBox();
            grpParts.Text = "1. Monte as partes";
            grpParts.Location = new Point(28, 410);
            grpParts.Size = new Size(665, 315);
            Controls.Add(grpParts);

            lvParts = new ListView();
            lvParts.Location = new Point(14, 24);
            lvParts.Size = new Size(505, 238);
            lvParts.View = View.Details;
            lvParts.FullRowSelect = true;
            lvParts.GridLines = true;
            lvParts.HideSelection = false;
            lvParts.MultiSelect = false;
            lvParts.Columns.Add("Parte", 65);
            lvParts.Columns.Add("Anexos", 57);
            lvParts.Columns.Add("Total", 75);
            lvParts.Columns.Add("E-mail est.", 82);
            lvParts.Columns.Add("Arquivos da parte", 220);
            grpParts.Controls.Add(lvParts);

            btnFilesToParts = NewButtonIn(grpParts, "Arquivos → partes", 530, 24, 120, 34);
            btnNewPart = NewButtonIn(grpParts, "+ Nova parte", 530, 65, 120, 34);
            btnDeletePart = NewButtonIn(grpParts, "Excluir parte", 530, 106, 120, 34);
            btnPartUp = NewButtonIn(grpParts, "↑ Subir", 530, 147, 57, 34);
            btnPartDown = NewButtonIn(grpParts, "↓ Descer", 593, 147, 57, 34);

            Label partsHelp = NewLabelIn(
                grpParts,
                "Atalho: “Arquivos → partes” cria uma parte para cada arquivo. Para uma parte com vários anexos, use “Nova parte”.",
                14, 269, 635, 38);
            partsHelp.ForeColor = SystemColors.GrayText;

            GroupBox grpAtt = new GroupBox();
            grpAtt.Text = "2. Anexos da parte selecionada";
            grpAtt.Location = new Point(704, 410);
            grpAtt.Size = new Size(438, 315);
            Controls.Add(grpAtt);

            lblSelectedPart = NewLabelIn(grpAtt, "Selecione uma parte", 14, 24, 300, 22);
            lblSelectedPart.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);

            lvAttachments = new ListView();
            lvAttachments.Location = new Point(14, 49);
            lvAttachments.Size = new Size(410, 167);
            lvAttachments.View = View.Details;
            lvAttachments.FullRowSelect = true;
            lvAttachments.GridLines = true;
            lvAttachments.HideSelection = false;
            lvAttachments.MultiSelect = true;
            lvAttachments.Columns.Add("Arquivo", 145);
            lvAttachments.Columns.Add("Tamanho", 78);
            lvAttachments.Columns.Add("Caminho", 180);
            grpAtt.Controls.Add(lvAttachments);

            btnAddAttachment = NewButtonIn(grpAtt, "+ Anexar à parte...", 14, 226, 135, 34);
            btnRemoveAttachment = NewButtonIn(grpAtt, "Remover anexo", 157, 226, 125, 34);
            btnSortAttachments = NewButtonIn(grpAtt, "Ordenar A-Z", 290, 226, 120, 34);

            Label attHelp = NewLabelIn(
                grpAtt,
                "Uma Parte pode ter 1 ou vários anexos. O mesmo arquivo pode aparecer em partes diferentes.",
                14, 267, 405, 36);
            attHelp.ForeColor = SystemColors.GrayText;

            lblCount = NewLabel("Partes: 0", 42, 734, 100, 22);
            lblPreview = NewLabel(
                "Monte as partes. O total e a numeração são calculados automaticamente.",
                145, 734, 997, 24);

            chkSizeWarning = new CheckBox();
            chkSizeWarning.Text = "Avisar se e-mail estimado exceder";
            chkSizeWarning.Location = new Point(42, 762);
            chkSizeWarning.Size = new Size(245, 26);
            chkSizeWarning.Checked = true;
            Controls.Add(chkSizeWarning);

            numMaxMB = new NumericUpDown();
            numMaxMB.Location = new Point(290, 762);
            numMaxMB.Size = new Size(64, 26);
            numMaxMB.Minimum = 1;
            numMaxMB.Maximum = 500;
            numMaxMB.Value = 20;
            Controls.Add(numMaxMB);

            NewLabel("MB", 360, 766, 35, 22);
            lblSizeWarning = NewLabel("Alerta: 20 MB por e-mail (estimativa).", 402, 765, 740, 24);

            chkSignature = new CheckBox();
            chkSignature.Text = "Incluir assinatura padrão do Outlook";
            chkSignature.Location = new Point(42, 794);
            chkSignature.Size = new Size(270, 26);
            chkSignature.Checked = true;
            Controls.Add(chkSignature);

            chkPartInBody = new CheckBox();
            chkPartInBody.Text = "Mostrar Parte 01/NN também no corpo";
            chkPartInBody.Location = new Point(320, 794);
            chkPartInBody.Size = new Size(270, 26);
            Controls.Add(chkPartInBody);

            rbDraft = new RadioButton();
            rbDraft.Text = "Criar em Rascunhos";
            rbDraft.Location = new Point(610, 794);
            rbDraft.Size = new Size(150, 26);
            rbDraft.Checked = true;
            Controls.Add(rbDraft);

            rbSend = new RadioButton();
            rbSend.Text = "Enviar";
            rbSend.Location = new Point(765, 794);
            rbSend.Size = new Size(80, 26);
            Controls.Add(rbSend);

            btnSave = NewButton("Salvar montagem", 856, 790, 135, 32);
            btnLoad = NewButton("Carregar montagem", 999, 790, 143, 32);

            btnReview = NewButton("REVISAR PARTE → ANEXOS", 42, 837, 255, 42);
            btnGenerate = NewButton("VALIDAR E GERAR E-MAILS", 310, 837, 270, 42);
            btnGenerate.Font = new Font("Segoe UI", 10.0f, FontStyle.Bold);

            lblStatus = NewLabel(
                "Modo seguro: os e-mails serão criados em Rascunhos.",
                600, 844, 542, 32);
        }

        private void WireEvents()
        {
            txtTo.TextChanged += delegate { MarkDirty(); };
            txtCC.TextChanged += delegate { MarkDirty(); };
            txtBCC.TextChanged += delegate { MarkDirty(); };
            txtSubject.TextChanged += delegate
            {
                MarkDirty();
                RefreshPartsView(GetSelectedPartIndex());
            };
            txtBody.TextChanged += delegate { MarkDirty(); };
            cmbAccount.SelectedIndexChanged += delegate { MarkDirty(); };
            chkSignature.CheckedChanged += delegate { MarkDirty(); };
            chkPartInBody.CheckedChanged += delegate { MarkDirty(); };
            chkSizeWarning.CheckedChanged += delegate { MarkDirty(); UpdateSizeWarning(); };
            numMaxMB.ValueChanged += delegate { MarkDirty(); UpdateSizeWarning(); };

            rbDraft.CheckedChanged += delegate
            {
                if (rbDraft.Checked) lblStatus.Text = "Modo seguro: os e-mails serão criados em Rascunhos.";
            };
            rbSend.CheckedChanged += delegate
            {
                if (rbSend.Checked) lblStatus.Text = "ATENÇÃO: modo Enviar. Haverá confirmação adicional.";
            };

            lvParts.SelectedIndexChanged += delegate
            {
                if (!_busy) RefreshAttachmentsView();
            };

            btnReload.Click += delegate
            {
                if (_busy) return;
                string keep = SelectedAccountKey();
                LoadAccounts(keep);
            };

            btnFilesToParts.Click += delegate { AddFilesAsParts(); };
            btnNewPart.Click += delegate { AddEmptyPart(); };
            btnDeletePart.Click += delegate { DeleteSelectedPart(); };
            btnPartUp.Click += delegate { MoveSelectedPart(-1); };
            btnPartDown.Click += delegate { MoveSelectedPart(1); };
            btnAddAttachment.Click += delegate { AddAttachmentsToSelectedPart(); };
            btnRemoveAttachment.Click += delegate { RemoveSelectedAttachments(); };
            btnSortAttachments.Click += delegate { SortSelectedPartAttachments(); };

            btnSave.Click += delegate
            {
                try { SaveSetup(); }
                catch (Exception ex) { ShowError("Não foi possível salvar a montagem.\r\n\r\n" + ex.Message); }
            };
            btnLoad.Click += delegate
            {
                try { LoadSetup(); }
                catch (Exception ex) { ShowError("Não foi possível carregar a montagem.\r\n\r\n" + ex.Message); }
            };
            btnReview.Click += delegate
            {
                try { ShowReview(); }
                catch (Exception ex) { ShowError(ex.Message); }
            };
            btnGenerate.Click += delegate
            {
                try { GenerateBatch(); }
                catch (OperationCanceledException) { SetBusy(false, "Operação cancelada."); }
                catch (Exception ex)
                {
                    SetBusy(false, "Operação interrompida.");
                    ShowError(ex.Message);
                }
            };

            FormClosing += OnFormClosing;
            FormClosed += OnFormClosed;
        }

        private Label NewLabel(string text, int x, int y, int width, int height)
        {
            Label l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.Size = new Size(width, height);
            Controls.Add(l);
            return l;
        }

        private Label NewLabelIn(Control parent, string text, int x, int y, int width, int height)
        {
            Label l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.Size = new Size(width, height);
            parent.Controls.Add(l);
            return l;
        }

        private TextBox NewTextBox(int x, int y, int width)
        {
            TextBox t = new TextBox();
            t.Location = new Point(x, y);
            t.Size = new Size(width, 27);
            Controls.Add(t);
            return t;
        }

        private Button NewButton(string text, int x, int y, int width, int height)
        {
            Button b = new Button();
            b.Text = text;
            b.Location = new Point(x, y);
            b.Size = new Size(width, height);
            Controls.Add(b);
            return b;
        }

        private Button NewButtonIn(Control parent, string text, int x, int y, int width, int height)
        {
            Button b = new Button();
            b.Text = text;
            b.Location = new Point(x, y);
            b.Size = new Size(width, height);
            parent.Controls.Add(b);
            return b;
        }

        private void MarkDirty()
        {
            if (!_loading && !_busy) _dirty = true;
        }

        private void CheckShutdown()
        {
            if (_hostShutdown || (_hostShuttingDown != null && _hostShuttingDown()))
                throw new OperationCanceledException("O Outlook está sendo encerrado.");
        }

        private void LoadAccounts(string preferredKey)
        {
            if (_busy) return;

            _loading = true;
            try
            {
                string oldKey = preferredKey;
                if (String.IsNullOrWhiteSpace(oldKey)) oldKey = SelectedAccountKey();

                ReleaseAccounts();
                cmbAccount.Items.Clear();

                dynamic session = null;
                dynamic accounts = null;
                try
                {
                    session = ((dynamic)_outlook).Session;
                    accounts = session.Accounts;
                    int count = (int)accounts.Count;

                    for (int i = 1; i <= count; i++)
                    {
                        object account = null;
                        try
                        {
                            account = accounts.Item(i);
                            dynamic a = account;

                            string smtp = "";
                            string display = "";
                            try { smtp = Convert.ToString(a.SmtpAddress); } catch { }
                            try { display = Convert.ToString(a.DisplayName); } catch { }

                            string key = !String.IsNullOrWhiteSpace(smtp) ? smtp : display;
                            string label = !String.IsNullOrWhiteSpace(smtp)
                                ? smtp + (!String.IsNullOrWhiteSpace(display) && !String.Equals(display, smtp, StringComparison.OrdinalIgnoreCase) ? " — " + display : "")
                                : display;

                            if (String.IsNullOrWhiteSpace(label)) label = "Conta " + i.ToString();
                            if (String.IsNullOrWhiteSpace(key)) key = label;

                            AccountInfo info = new AccountInfo();
                            info.ComObject = account;
                            info.Display = label;
                            info.Key = key;
                            _accounts.Add(info);
                            cmbAccount.Items.Add(info);
                            account = null;
                        }
                        finally
                        {
                            if (account != null) ReleaseCom(account);
                        }
                    }
                }
                finally
                {
                    ReleaseCom(accounts);
                    ReleaseCom(session);
                }

                if (cmbAccount.Items.Count > 0)
                {
                    int selected = 0;
                    if (!String.IsNullOrWhiteSpace(oldKey))
                    {
                        for (int i = 0; i < _accounts.Count; i++)
                        {
                            if (String.Equals(_accounts[i].Key, oldKey, StringComparison.OrdinalIgnoreCase))
                            {
                                selected = i;
                                break;
                            }
                        }
                    }
                    cmbAccount.SelectedIndex = selected;
                }

                lblStatus.Text = cmbAccount.Items.Count > 0
                    ? "Outlook conectado. Modo seguro: Rascunhos."
                    : "Nenhuma conta do Outlook foi encontrada.";
            }
            catch (Exception ex)
            {
                Connect.Log("LoadAccounts ERROR: " + ex.ToString());
                ShowError("Não foi possível listar as contas do Outlook.\r\n\r\n" + ex.Message);
            }
            finally
            {
                _loading = false;
            }
        }

        private string SelectedAccountKey()
        {
            AccountInfo info = cmbAccount.SelectedItem as AccountInfo;
            return info == null ? "" : (info.Key ?? "");
        }

        private AccountInfo SelectedAccount()
        {
            return cmbAccount.SelectedItem as AccountInfo;
        }

        private int GetSelectedPartIndex()
        {
            if (lvParts.SelectedIndices.Count != 1) return -1;
            return lvParts.SelectedIndices[0];
        }

        private void SelectPart(int index)
        {
            if (index < 0 || index >= lvParts.Items.Count) return;
            foreach (ListViewItem item in lvParts.Items) item.Selected = false;
            lvParts.Items[index].Selected = true;
            lvParts.Items[index].Focused = true;
            lvParts.EnsureVisible(index);
        }

        private void RefreshPartsView(int selectIndex)
        {
            int old = GetSelectedPartIndex();
            if (selectIndex < -1) selectIndex = old;

            lvParts.BeginUpdate();
            try
            {
                lvParts.Items.Clear();
                int total = _parts.Count;
                int width = Math.Max(2, total.ToString().Length);

                for (int i = 0; i < total; i++)
                {
                    PartModel p = _parts[i];

                    ListViewItem row = new ListViewItem(
                        (i + 1).ToString("D" + width.ToString()) + "/" +
                        total.ToString("D" + width.ToString()));

                    row.SubItems.Add(p.Attachments.Count.ToString());
                    row.SubItems.Add(FormatBytes(GetPartRawBytes(p)));
                    row.SubItems.Add(FormatBytes(EstimatePartMimeBytes(p)));
                    row.SubItems.Add(GetFilesSummary(p));
                    lvParts.Items.Add(row);
                }
            }
            finally
            {
                lvParts.EndUpdate();
            }

            lblCount.Text = "Partes: " + _parts.Count.ToString();

            if (_parts.Count > 0 && !String.IsNullOrWhiteSpace(txtSubject.Text))
            {
                int width = Math.Max(2, _parts.Count.ToString().Length);
                string first = "Parte " + 1.ToString("D" + width.ToString()) + "/" + _parts.Count.ToString("D" + width.ToString());
                string last = "Parte " + _parts.Count.ToString("D" + width.ToString()) + "/" + _parts.Count.ToString("D" + width.ToString());
                lblPreview.Text = "Assuntos: " + txtSubject.Text.Trim() + " - " + first + "   …   " + last;
            }
            else
            {
                lblPreview.Text = "Monte as partes. O total e a numeração são calculados automaticamente.";
            }

            UpdateSizeWarning();

            if (_parts.Count > 0)
            {
                if (selectIndex < 0) selectIndex = old >= 0 ? old : 0;
                if (selectIndex >= _parts.Count) selectIndex = _parts.Count - 1;
                SelectPart(selectIndex);
            }
            else
            {
                RefreshAttachmentsView();
            }
        }

        private void RefreshAttachmentsView()
        {
            lvAttachments.Items.Clear();
            int idx = GetSelectedPartIndex();
            if (idx < 0 || idx >= _parts.Count)
            {
                lblSelectedPart.Text = "Selecione uma parte";
                return;
            }

            int total = _parts.Count;
            int width = Math.Max(2, total.ToString().Length);
            lblSelectedPart.Text =
                "Anexos da Parte " +
                (idx + 1).ToString("D" + width.ToString()) + "/" +
                total.ToString("D" + width.ToString());

            foreach (string path in _parts[idx].Attachments)
            {
                ListViewItem row = new ListViewItem(Path.GetFileName(path));
                if (File.Exists(path))
                {
                    FileInfo fi = new FileInfo(path);
                    row.SubItems.Add(FormatBytes(fi.Length));
                    row.SubItems.Add(fi.FullName);
                }
                else
                {
                    row.SubItems.Add("AUSENTE");
                    row.SubItems.Add(path);
                }
                row.Tag = path;
                lvAttachments.Items.Add(row);
            }
        }

        private void AddFilesAsParts()
        {
            if (_busy) return;

            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Multiselect = true;
                dlg.Title = "Selecione os arquivos — será criada uma Parte para cada arquivo";
                dlg.Filter = "Todos os arquivos (*.*)|*.*";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                List<string> files = new List<string>(dlg.FileNames);
                files.Sort(NaturalPathCompare);

                int firstNew = _parts.Count;
                foreach (string file in files)
                {
                    PartModel p = new PartModel();
                    p.Attachments.Add(Path.GetFullPath(file));
                    _parts.Add(p);
                }
                _dirty = true;
                RefreshPartsView(firstNew);
            }
        }

        private void AddEmptyPart()
        {
            if (_busy) return;
            _parts.Add(new PartModel());
            _dirty = true;
            RefreshPartsView(_parts.Count - 1);
        }

        private void DeleteSelectedPart()
        {
            if (_busy) return;
            int idx = GetSelectedPartIndex();
            if (idx < 0)
            {
                ShowInfo("Selecione uma parte para excluir.");
                return;
            }

            string msg = "Excluir a Parte " + (idx + 1).ToString() + "?";
            if (_parts[idx].Attachments.Count > 0)
                msg += "\r\n\r\nA associação com " + _parts[idx].Attachments.Count.ToString() + " anexo(s) será removida.";
            msg += "\r\n\r\nOs arquivos originais não serão apagados.";

            if (MessageBox.Show(this, msg, "Excluir parte", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            _parts.RemoveAt(idx);
            _dirty = true;
            RefreshPartsView(Math.Min(idx, _parts.Count - 1));
        }

        private void MoveSelectedPart(int delta)
        {
            if (_busy) return;
            int idx = GetSelectedPartIndex();
            if (idx < 0)
            {
                ShowInfo("Selecione uma parte para mover.");
                return;
            }

            int target = idx + delta;
            if (target < 0 || target >= _parts.Count) return;

            PartModel p = _parts[idx];
            _parts.RemoveAt(idx);
            _parts.Insert(target, p);
            _dirty = true;
            RefreshPartsView(target);
        }

        private void AddAttachmentsToSelectedPart()
        {
            if (_busy) return;
            int idx = GetSelectedPartIndex();
            if (idx < 0)
            {
                ShowInfo("Primeiro selecione ou crie uma Parte.");
                return;
            }

            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Multiselect = true;
                dlg.Title = "Selecione um ou vários anexos para esta Parte";
                dlg.Filter = "Todos os arquivos (*.*)|*.*";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                int added = 0;
                foreach (string file in dlg.FileNames)
                {
                    string full = Path.GetFullPath(file);
                    if (!ContainsPath(_parts[idx].Attachments, full))
                    {
                        _parts[idx].Attachments.Add(full);
                        added++;
                    }
                }

                _dirty = true;
                RefreshPartsView(idx);

                if (added < dlg.FileNames.Length)
                    ShowInfo((dlg.FileNames.Length - added).ToString() + " arquivo(s) já estavam nesta Parte.");
            }
        }

        private void RemoveSelectedAttachments()
        {
            if (_busy) return;
            int idx = GetSelectedPartIndex();
            if (idx < 0) return;

            List<string> selected = new List<string>();
            foreach (ListViewItem item in lvAttachments.SelectedItems)
            {
                string path = item.Tag as string;
                if (!String.IsNullOrWhiteSpace(path)) selected.Add(path);
            }

            if (selected.Count == 0)
            {
                ShowInfo("Selecione um ou mais anexos para remover desta Parte.");
                return;
            }

            for (int i = _parts[idx].Attachments.Count - 1; i >= 0; i--)
            {
                if (ContainsPath(selected, _parts[idx].Attachments[i]))
                    _parts[idx].Attachments.RemoveAt(i);
            }

            _dirty = true;
            RefreshPartsView(idx);
        }

        private void SortSelectedPartAttachments()
        {
            if (_busy) return;
            int idx = GetSelectedPartIndex();
            if (idx < 0) return;

            _parts[idx].Attachments.Sort(NaturalPathCompare);
            _dirty = true;
            RefreshPartsView(idx);
        }

        private void UpdateSizeWarning()
        {
            if (lblSizeWarning == null) return;

            if (!chkSizeWarning.Checked)
            {
                lblSizeWarning.Text = "Alerta de tamanho desativado.";
                return;
            }

            long limit = (long)numMaxMB.Value * 1024L * 1024L;
            List<int> overs = new List<int>();
            for (int i = 0; i < _parts.Count; i++)
            {
                if (EstimatePartMimeBytes(_parts[i]) > limit) overs.Add(i + 1);
            }

            if (overs.Count == 0)
            {
                lblSizeWarning.Text =
                    "Todas as partes estão dentro do alerta configurado de " +
                    ((int)numMaxMB.Value).ToString() + " MB.";
            }
            else
            {
                lblSizeWarning.Text = "ATENÇÃO: parte(s) " + JoinInts(overs) + " excedem o alerta de tamanho.";
            }
        }

        private BatchSnapshot BuildSnapshot()
        {
            BatchSnapshot s = new BatchSnapshot();
            s.Account = SelectedAccount();
            s.To = NormalizeAddresses(txtTo.Text);
            s.Cc = NormalizeAddresses(txtCC.Text);
            s.Bcc = NormalizeAddresses(txtBCC.Text);
            s.Subject = txtSubject.Text.Trim();
            s.Body = txtBody.Text;
            s.IncludeSignature = chkSignature.Checked;
            s.PartInBody = chkPartInBody.Checked;
            s.SendNow = rbSend.Checked;
            s.SizeWarning = chkSizeWarning.Checked;
            s.SizeWarningMb = (int)numMaxMB.Value;
            s.Parts = new List<PartSnapshot>();

            foreach (PartModel p in _parts)
            {
                PartSnapshot ps = new PartSnapshot();
                ps.Attachments = p.Attachments.ToArray();
                s.Parts.Add(ps);
            }
            return s;
        }

        private void ValidateSnapshot(BatchSnapshot s, bool askSizeWarning)
        {
            CheckShutdown();

            if (s.Account == null)
                throw new InvalidOperationException("Selecione uma conta de envio.");

            if (String.IsNullOrWhiteSpace(s.To))
                throw new InvalidOperationException("Preencha pelo menos um destinatário no campo Para.");

            if (String.IsNullOrWhiteSpace(s.Subject))
                throw new InvalidOperationException("Preencha o Assunto base.");

            if (s.Parts == null || s.Parts.Count < 1)
                throw new InvalidOperationException("Crie pelo menos uma parte.");

            int total = s.Parts.Count;
            int width = Math.Max(2, total.ToString().Length);
            string longestSubject =
                s.Subject + " - Parte " +
                total.ToString("D" + width.ToString()) + "/" +
                total.ToString("D" + width.ToString());

            if (longestSubject.Length > 240)
                throw new InvalidOperationException(
                    "O assunto final ficaria muito longo (" + longestSubject.Length.ToString() +
                    " caracteres). Reduza o Assunto base.");

            for (int i = 0; i < s.Parts.Count; i++)
            {
                string[] files = s.Parts[i].Attachments;
                if (files == null || files.Length == 0)
                    throw new InvalidOperationException(
                        "A Parte " + (i + 1).ToString() +
                        " está sem anexos. Adicione pelo menos um arquivo ou remova a parte.");

                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string path in files)
                {
                    if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
                        throw new FileNotFoundException(
                            "Arquivo não encontrado na Parte " + (i + 1).ToString() + ":\r\n" + path);

                    string full = Path.GetFullPath(path);
                    if (!seen.Add(full))
                        throw new InvalidOperationException(
                            "Arquivo duplicado dentro da Parte " + (i + 1).ToString() + ":\r\n" + full);
                }
            }

            ValidateRecipients(s);

            if (askSizeWarning && s.SizeWarning)
            {
                long limit = (long)s.SizeWarningMb * 1024L * 1024L;
                List<string> overs = new List<string>();

                for (int i = 0; i < s.Parts.Count; i++)
                {
                    long est = EstimateSnapshotMimeBytes(s.Parts[i]);
                    if (est > limit)
                    {
                        overs.Add(
                            "Parte " + (i + 1).ToString() +
                            ": " + s.Parts[i].Attachments.Length.ToString() +
                            " anexo(s), estimado " + FormatBytes(est));
                    }
                }

                if (overs.Count > 0)
                {
                    StringBuilder b = new StringBuilder();
                    b.Append("Uma ou mais partes ultrapassam o limite de ALERTA de ");
                    b.Append(s.SizeWarningMb);
                    b.Append(" MB.\r\n");
                    b.Append("Este valor é apenas uma estimativa MIME; o servidor define o limite real.\r\n\r\n");

                    int max = Math.Min(12, overs.Count);
                    for (int i = 0; i < max; i++) b.AppendLine(overs[i]);
                    if (overs.Count > max) b.AppendLine("... e mais " + (overs.Count - max).ToString() + ".");

                    b.Append("\r\nDeseja continuar mesmo assim?");

                    if (MessageBox.Show(
                        this, b.ToString(), "Aviso de tamanho",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                        throw new OperationCanceledException();
                }
            }
        }

        private void ValidateRecipients(BatchSnapshot s)
        {
            object mail = null;
            try
            {
                mail = NewMailItem(s.Account);
                dynamic m = mail;
                m.To = s.To;
                if (!String.IsNullOrWhiteSpace(s.Cc)) m.CC = s.Cc;
                if (!String.IsNullOrWhiteSpace(s.Bcc)) m.BCC = s.Bcc;

                dynamic recipients = null;
                try
                {
                    recipients = m.Recipients;
                    bool ok = (bool)recipients.ResolveAll();
                    if (!ok)
                    {
                        List<string> bad = new List<string>();
                        int count = (int)recipients.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            object rObj = null;
                            try
                            {
                                rObj = recipients.Item(i);
                                dynamic r = rObj;
                                bool resolved = false;
                                try { resolved = (bool)r.Resolved; } catch { }
                                if (!resolved)
                                {
                                    string name = "";
                                    try { name = Convert.ToString(r.Name); } catch { }
                                    if (!String.IsNullOrWhiteSpace(name)) bad.Add(name);
                                }
                            }
                            finally
                            {
                                ReleaseCom(rObj);
                            }
                        }

                        throw new InvalidOperationException(
                            bad.Count > 0
                            ? "Destinatários não resolvidos: " + String.Join("; ", bad.ToArray())
                            : "O Outlook não conseguiu resolver um ou mais destinatários.");
                    }
                }
                finally
                {
                    ReleaseCom(recipients);
                }
            }
            finally
            {
                if (mail != null)
                {
                    try { ((dynamic)mail).Close(1); } catch { }
                    ReleaseCom(mail);
                }
            }
        }

        private object NewMailItem(AccountInfo account)
        {
            CheckShutdown();

            object folder = null;
            object store = null;
            object items = null;
            object session = null;

            try
            {
                if (account != null && account.ComObject != null)
                {
                    try
                    {
                        dynamic a = account.ComObject;
                        store = a.DeliveryStore;
                        if (store != null)
                        {
                            folder = ((dynamic)store).GetDefaultFolder(16);
                        }
                    }
                    catch
                    {
                        ReleaseCom(folder); folder = null;
                        ReleaseCom(store); store = null;
                    }
                }

                if (folder == null)
                {
                    session = ((dynamic)_outlook).Session;
                    folder = ((dynamic)session).GetDefaultFolder(16);
                }

                items = ((dynamic)folder).Items;
                object mail = ((dynamic)items).Add("IPM.Note");

                if (account != null && account.ComObject != null)
                {
                    try
                    {
                        ((dynamic)mail).SendUsingAccount = account.ComObject;
                    }
                    catch
                    {
                        try { ((dynamic)mail).Close(1); } catch { }
                        ReleaseCom(mail);
                        throw new InvalidOperationException(
                            "Não foi possível definir a conta de envio selecionada.");
                    }
                }

                return mail;
            }
            finally
            {
                ReleaseCom(items);
                ReleaseCom(folder);
                ReleaseCom(store);
                ReleaseCom(session);
            }
        }

        private void ShowReview()
        {
            BatchSnapshot s = BuildSnapshot();
            ValidateSnapshot(s, false);

            StringBuilder b = new StringBuilder();
            b.AppendLine("OUTLOOK PART SENDER — REVISÃO DO LOTE");
            b.AppendLine(new String('=', 72));
            b.AppendLine("Conta: " + s.Account.Display);
            b.AppendLine("Para: " + s.To);
            if (!String.IsNullOrWhiteSpace(s.Cc)) b.AppendLine("CC: " + s.Cc);
            if (!String.IsNullOrWhiteSpace(s.Bcc)) b.AppendLine("BCC: " + s.Bcc);
            b.AppendLine("Modo: " + (s.SendNow ? "ENVIAR" : "RASCUNHOS"));
            b.AppendLine("Partes: " + s.Parts.Count.ToString());
            b.AppendLine();

            int width = Math.Max(2, s.Parts.Count.ToString().Length);
            for (int i = 0; i < s.Parts.Count; i++)
            {
                string label =
                    "Parte " +
                    (i + 1).ToString("D" + width.ToString()) + "/" +
                    s.Parts.Count.ToString("D" + width.ToString());

                b.AppendLine(label);
                b.AppendLine("Assunto: " + s.Subject + " - " + label);
                b.AppendLine("Anexos: " + s.Parts[i].Attachments.Length.ToString());
                b.AppendLine("Estimado: " + FormatBytes(EstimateSnapshotMimeBytes(s.Parts[i])));

                foreach (string path in s.Parts[i].Attachments)
                {
                    b.AppendLine("  • " + Path.GetFileName(path));
                    b.AppendLine("    " + path);
                }
                b.AppendLine();
            }

            using (ReviewForm review = new ReviewForm(b.ToString()))
            {
                review.ShowDialog(this);
            }
        }

        private void GenerateBatch()
        {
            CheckShutdown();

            BatchSnapshot s = BuildSnapshot();
            ValidateSnapshot(s, true);

            string confirm = BuildConfirmation(s);
            if (MessageBox.Show(
                this, confirm, "Confirmar geração",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            CheckShutdown();

            if (s.SendNow)
            {
                if (MessageBox.Show(
                    this,
                    "ATENÇÃO: os e-mails serão entregues ao Outlook para envio.\r\n\r\n" +
                    "Serão " + s.Parts.Count.ToString() + " mensagens.\r\n\r\n" +
                    "Confirma o envio automático?",
                    "CONFIRMAR ENVIO AUTOMÁTICO",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning) != DialogResult.Yes)
                    return;
            }

            CheckShutdown();

            int total = s.Parts.Count;
            int width = Math.Max(2, total.ToString().Length);
            string batchId = Guid.NewGuid().ToString();
            int completed = 0;

            SetBusy(true, "Preparando lote...");

            try
            {
                for (int i = 0; i < total; i++)
                {
                    CheckShutdown();

                    string partLabel =
                        "Parte " +
                        (i + 1).ToString("D" + width.ToString()) + "/" +
                        total.ToString("D" + width.ToString());

                    SetBusy(
                        true,
                        "Criando " + (i + 1).ToString() + "/" + total.ToString() +
                        " — " + partLabel +
                        " — " + s.Parts[i].Attachments.Length.ToString() + " anexo(s)");

                    Application.DoEvents();
                    CheckShutdown();

                    object mail = null;
                    bool committed = false;

                    try
                    {
                        mail = NewMailItem(s.Account);
                        dynamic m = mail;

                        m.To = s.To;
                        if (!String.IsNullOrWhiteSpace(s.Cc)) m.CC = s.Cc;
                        if (!String.IsNullOrWhiteSpace(s.Bcc)) m.BCC = s.Bcc;
                        m.Subject = s.Subject + " - " + partLabel;

                        TrySetBatchProperty(mail, batchId);

                        if (s.IncludeSignature)
                        {
                            InsertBodyUsingWordEditor(mail, s.Body, partLabel, s.PartInBody);
                        }
                        else
                        {
                            m.HTMLBody = BuildHtmlBody(s.Body, partLabel, s.PartInBody);
                        }

                        dynamic attachments = null;
                        try
                        {
                            attachments = m.Attachments;
                            foreach (string path in s.Parts[i].Attachments)
                            {
                                attachments.Add(path);
                            }
                        }
                        finally
                        {
                            ReleaseCom(attachments);
                        }

                        dynamic recipients = null;
                        try
                        {
                            recipients = m.Recipients;
                            if (!(bool)recipients.ResolveAll())
                                throw new InvalidOperationException(
                                    "O Outlook deixou de conseguir resolver um ou mais destinatários na " +
                                    partLabel + ".");
                        }
                        finally
                        {
                            ReleaseCom(recipients);
                        }

                        CheckShutdown();

                        if (s.SendNow)
                        {
                            m.Send();
                            committed = true;
                        }
                        else
                        {
                            m.Save();
                            committed = true;
                            if (s.IncludeSignature)
                            {
                                try { m.Close(0); } catch { }
                            }
                        }

                        completed++;
                    }
                    catch
                    {
                        if (mail != null && !committed)
                        {
                            try { ((dynamic)mail).Close(1); } catch { }
                        }
                        throw;
                    }
                    finally
                    {
                        ReleaseCom(mail);
                    }
                }

                if (s.SendNow)
                {
                    lblStatus.Text =
                        "Concluído: " + completed.ToString() +
                        " mensagens entregues ao Outlook para envio.";
                    ShowInfo(
                        completed.ToString() +
                        " mensagens foram entregues ao Outlook para envio.\r\n\r\n" +
                        "Isso não confirma recebimento pelo destinatário. Verifique Caixa de Saída, " +
                        "Itens Enviados e eventuais rejeições.");
                }
                else
                {
                    lblStatus.Text =
                        "Concluído: " + completed.ToString() + " rascunhos criados.";
                    ShowInfo(
                        completed.ToString() +
                        " e-mails foram criados em Rascunhos.\r\n\r\n" +
                        "Cada rascunho contém somente os anexos configurados para sua Parte. " +
                        "Revise antes de enviar.");
                }
            }
            catch (Exception ex)
            {
                Connect.Log(
                    "Batch ERROR id=" + batchId +
                    " completed=" + completed.ToString() + "/" + total.ToString() +
                    " " + ex.ToString());

                throw new InvalidOperationException(
                    "Operação interrompida.\r\n\r\n" +
                    completed.ToString() + " de " + total.ToString() +
                    (s.SendNow ? " mensagens foram entregues ao Outlook antes do erro." : " rascunhos foram criados antes do erro.") +
                    "\r\n\r\nErro: " + ex.Message +
                    "\r\n\r\nLote: " + batchId +
                    "\r\n\r\nConfira Rascunhos/Caixa de Saída antes de executar novamente para evitar duplicidade.",
                    ex);
            }
            finally
            {
                if (!_hostShutdown)
                {
                    rbDraft.Checked = true;
                    rbSend.Checked = false;
                    SetBusy(false, "Modo seguro: os e-mails serão criados em Rascunhos.");
                }
            }
        }

        private void InsertBodyUsingWordEditor(object mail, string body, string partLabel, bool partInBody)
        {
            dynamic m = mail;
            m.Display(false);

            for (int i = 0; i < 6; i++)
            {
                Application.DoEvents();
                CheckShutdown();
                Thread.Sleep(100);
            }

            object inspector = null;
            object doc = null;
            object range = null;

            try
            {
                inspector = m.GetInspector;
                doc = ((dynamic)inspector).WordEditor;
                if (doc == null)
                    throw new InvalidOperationException(
                        "O editor do Outlook não ficou disponível para inserir a mensagem/assinatura.");

                string prefix = "";
                if (partInBody) prefix += partLabel + "\r\n\r\n";
                prefix += body ?? "";
                if (!String.IsNullOrWhiteSpace(prefix)) prefix += "\r\n\r\n";

                range = ((dynamic)doc).Range(0, 0);
                ((dynamic)range).InsertBefore(prefix);
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(doc);
                ReleaseCom(inspector);
            }
        }

        private void TrySetBatchProperty(object mail, string batchId)
        {
            object props = null;
            object prop = null;
            try
            {
                props = ((dynamic)mail).UserProperties;
                prop = ((dynamic)props).Add("OutlookPartSenderBatchId", 1, false);
                ((dynamic)prop).Value = batchId;
            }
            catch
            {
            }
            finally
            {
                ReleaseCom(prop);
                ReleaseCom(props);
            }
        }

        private string BuildConfirmation(BatchSnapshot s)
        {
            int totalAttachments = 0;
            foreach (PartSnapshot p in s.Parts) totalAttachments += p.Attachments.Length;

            int width = Math.Max(2, s.Parts.Count.ToString().Length);
            string first = "Parte " + 1.ToString("D" + width.ToString()) + "/" + s.Parts.Count.ToString("D" + width.ToString());
            string last = "Parte " + s.Parts.Count.ToString("D" + width.ToString()) + "/" + s.Parts.Count.ToString("D" + width.ToString());

            StringBuilder b = new StringBuilder();
            b.AppendLine("Conta: " + s.Account.Display);
            b.AppendLine("Para: " + s.To);
            if (!String.IsNullOrWhiteSpace(s.Cc)) b.AppendLine("CC: " + s.Cc);
            if (!String.IsNullOrWhiteSpace(s.Bcc)) b.AppendLine("BCC: " + s.Bcc);
            b.AppendLine("Modo: " + (s.SendNow ? "ENTREGAR AO OUTLOOK PARA ENVIO" : "CRIAR EM RASCUNHOS"));
            b.AppendLine("E-mails / partes: " + s.Parts.Count.ToString());
            b.AppendLine("Anexos totais: " + totalAttachments.ToString());
            b.AppendLine();
            b.AppendLine("Primeiro: " + s.Subject + " - " + first);
            b.AppendLine("Último: " + s.Subject + " - " + last);
            b.AppendLine();
            b.AppendLine("Cada e-mail receberá exatamente os anexos configurados para aquela Parte.");
            b.AppendLine();
            b.Append("Continuar?");
            return b.ToString();
        }

        private void SetBusy(bool busy, string status)
        {
            _busy = busy;

            cmbAccount.Enabled = !busy;
            txtTo.Enabled = !busy;
            txtCC.Enabled = !busy;
            txtBCC.Enabled = !busy;
            txtSubject.Enabled = !busy;
            txtBody.Enabled = !busy;
            lvParts.Enabled = !busy;
            lvAttachments.Enabled = !busy;
            chkSignature.Enabled = !busy;
            chkPartInBody.Enabled = !busy;
            chkSizeWarning.Enabled = !busy;
            numMaxMB.Enabled = !busy;
            rbDraft.Enabled = !busy;
            rbSend.Enabled = !busy;

            Button[] buttons = new Button[]
            {
                btnReload, btnFilesToParts, btnNewPart, btnDeletePart,
                btnPartUp, btnPartDown, btnAddAttachment, btnRemoveAttachment,
                btnSortAttachments, btnSave, btnLoad, btnReview, btnGenerate
            };
            foreach (Button b in buttons) b.Enabled = !busy;

            UseWaitCursor = busy;
            if (!String.IsNullOrWhiteSpace(status)) lblStatus.Text = status;
            Application.DoEvents();
        }

        private void SaveSetup()
        {
            if (_busy) return;

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Title = "Salvar montagem do Outlook Part Sender";
                dlg.Filter = "Outlook Part Sender (*.opsparts.json)|*.opsparts.json|JSON (*.json)|*.json";
                dlg.FileName = "montagem-email-partes.opsparts.json";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                SetupConfig cfg = new SetupConfig();
                cfg.Version = "3.1.0";
                cfg.AccountKey = SelectedAccountKey();
                cfg.To = txtTo.Text;
                cfg.CC = txtCC.Text;
                cfg.BCC = txtBCC.Text;
                cfg.Subject = txtSubject.Text;
                cfg.Body = txtBody.Text;
                cfg.IncludeSignature = chkSignature.Checked;
                cfg.PartInBody = chkPartInBody.Checked;
                cfg.SizeWarning = chkSizeWarning.Checked;
                cfg.SizeWarningMB = (int)numMaxMB.Value;
                cfg.Parts = new List<SetupPart>();

                foreach (PartModel p in _parts)
                {
                    SetupPart sp = new SetupPart();
                    sp.Attachments = new List<string>(p.Attachments);
                    cfg.Parts.Add(sp);
                }

                JavaScriptSerializer ser = new JavaScriptSerializer();
                ser.MaxJsonLength = Int32.MaxValue;
                string json = ser.Serialize(cfg);
                File.WriteAllText(dlg.FileName, json, new UTF8Encoding(false));

                _dirty = false;
                ShowInfo("Montagem salva.\r\n\r\n" + dlg.FileName);
            }
        }

        private void LoadSetup()
        {
            if (_busy) return;

            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Carregar montagem do Outlook Part Sender";
                dlg.Filter = "Outlook Part Sender (*.opsparts.json)|*.opsparts.json|JSON (*.json)|*.json";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                string raw = File.ReadAllText(dlg.FileName, Encoding.UTF8);
                JavaScriptSerializer ser = new JavaScriptSerializer();
                ser.MaxJsonLength = Int32.MaxValue;
                SetupConfig cfg = ser.Deserialize<SetupConfig>(raw);
                if (cfg == null) throw new InvalidDataException("Arquivo de montagem inválido.");

                _loading = true;
                try
                {
                    txtTo.Text = cfg.To ?? "";
                    txtCC.Text = cfg.CC ?? "";
                    txtBCC.Text = cfg.BCC ?? "";
                    txtSubject.Text = cfg.Subject ?? "";
                    txtBody.Text = cfg.Body ?? "";
                    chkSignature.Checked = cfg.IncludeSignature;
                    chkPartInBody.Checked = cfg.PartInBody;
                    chkSizeWarning.Checked = cfg.SizeWarning;

                    int mb = cfg.SizeWarningMB <= 0 ? 20 : cfg.SizeWarningMB;
                    mb = Math.Max((int)numMaxMB.Minimum, Math.Min((int)numMaxMB.Maximum, mb));
                    numMaxMB.Value = mb;

                    _parts.Clear();
                    if (cfg.Parts != null)
                    {
                        foreach (SetupPart sp in cfg.Parts)
                        {
                            PartModel p = new PartModel();
                            if (sp != null && sp.Attachments != null)
                                p.Attachments.AddRange(sp.Attachments);
                            _parts.Add(p);
                        }
                    }

                    if (!String.IsNullOrWhiteSpace(cfg.AccountKey))
                    {
                        for (int i = 0; i < _accounts.Count; i++)
                        {
                            if (String.Equals(_accounts[i].Key, cfg.AccountKey, StringComparison.OrdinalIgnoreCase))
                            {
                                cmbAccount.SelectedIndex = i;
                                break;
                            }
                        }
                    }

                    rbDraft.Checked = true;
                    rbSend.Checked = false;
                    RefreshPartsView(_parts.Count > 0 ? 0 : -1);
                }
                finally
                {
                    _loading = false;
                }

                _dirty = false;

                List<string> missing = new List<string>();
                foreach (PartModel p in _parts)
                    foreach (string path in p.Attachments)
                        if (!File.Exists(path)) missing.Add(path);

                if (missing.Count > 0)
                {
                    ShowInfo(
                        "Montagem carregada, mas " + missing.Count.ToString() +
                        " arquivo(s) não foram encontrados.\r\n\r\n" +
                        "Eles aparecem como AUSENTES e precisam ser corrigidos antes da geração.");
                }
                else
                {
                    ShowInfo("Montagem carregada com sucesso.\r\n\r\nPor segurança, o modo voltou para Rascunhos.");
                }
            }
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (_hostShutdown || (_hostShuttingDown != null && _hostShuttingDown()))
            {
                return;
            }

            if (_busy)
            {
                e.Cancel = true;
                ShowInfo("A geração está em andamento. Aguarde terminar antes de fechar.");
                return;
            }

            if (_dirty)
            {
                DialogResult r = MessageBox.Show(
                    this,
                    "Existem alterações de montagem não salvas.\r\n\r\nFechar mesmo assim?",
                    "Outlook Part Sender",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (r != DialogResult.Yes) e.Cancel = true;
            }
        }

        private void OnFormClosed(object sender, FormClosedEventArgs e)
        {
            ReleaseAccounts();
        }

        private void ReleaseAccounts()
        {
            foreach (AccountInfo a in _accounts)
            {
                ReleaseCom(a.ComObject);
                a.ComObject = null;
            }
            _accounts.Clear();
        }

        private static bool ContainsPath(IList<string> paths, string target)
        {
            foreach (string p in paths)
                if (String.Equals(Path.GetFullPath(p), Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private static int NaturalPathCompare(string x, string y)
        {
            string a = NaturalKey(Path.GetFileName(x));
            string b = NaturalKey(Path.GetFileName(y));
            return StringComparer.CurrentCultureIgnoreCase.Compare(a, b);
        }

        private static string NaturalKey(string value)
        {
            if (value == null) return "";
            return Regex.Replace(
                value.ToLowerInvariant(),
                "\\d+",
                delegate(Match m) { return m.Value.PadLeft(20, '0'); });
        }

        private static string NormalizeAddresses(string text)
        {
            if (String.IsNullOrWhiteSpace(text)) return "";
            string[] parts = text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> clean = new List<string>();
            foreach (string p in parts)
            {
                string v = p.Trim();
                if (!String.IsNullOrWhiteSpace(v)) clean.Add(v);
            }
            return String.Join("; ", clean.ToArray());
        }

        private static long GetPartRawBytes(PartModel p)
        {
            long total = 0;
            foreach (string path in p.Attachments)
                if (File.Exists(path)) total += new FileInfo(path).Length;
            return total;
        }

        private static long EstimatePartMimeBytes(PartModel p)
        {
            long total = 512L * 1024L;
            foreach (string path in p.Attachments)
            {
                if (!File.Exists(path)) continue;
                long size = new FileInfo(path).Length;
                total += ((size + 2L) / 3L) * 4L;
                total += 16L * 1024L;
            }
            return total;
        }

        private static long EstimateSnapshotMimeBytes(PartSnapshot p)
        {
            long total = 512L * 1024L;
            foreach (string path in p.Attachments)
            {
                if (!File.Exists(path)) continue;
                long size = new FileInfo(path).Length;
                total += ((size + 2L) / 3L) * 4L;
                total += 16L * 1024L;
            }
            return total;
        }

        private static string GetFilesSummary(PartModel p)
        {
            if (p.Attachments.Count == 0) return "(sem anexos)";

            List<string> names = new List<string>();
            foreach (string path in p.Attachments) names.Add(Path.GetFileName(path));
            string text = String.Join("; ", names.ToArray());
            return text.Length > 105 ? text.Substring(0, 102) + "..." : text;
        }

        private static string FormatBytes(long bytes)
        {
            if (bytes >= 1024L * 1024L * 1024L)
                return (bytes / (1024.0 * 1024.0 * 1024.0)).ToString("N2") + " GB";
            if (bytes >= 1024L * 1024L)
                return (bytes / (1024.0 * 1024.0)).ToString("N2") + " MB";
            if (bytes >= 1024L)
                return (bytes / 1024.0).ToString("N1") + " KB";
            return bytes.ToString() + " B";
        }

        private static string JoinInts(List<int> values)
        {
            List<string> strings = new List<string>();
            foreach (int v in values) strings.Add(v.ToString());
            return String.Join(", ", strings.ToArray());
        }

        private static string BuildHtmlBody(string body, string partLabel, bool partInBody)
        {
            string encoded = WebUtility.HtmlEncode(body ?? "")
                .Replace("\r\n", "<br>")
                .Replace("\n", "<br>");

            StringBuilder b = new StringBuilder();
            b.Append("<html><body><div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt;\">");
            if (partInBody)
            {
                b.Append("<div style=\"margin-bottom:10px;\"><b>");
                b.Append(WebUtility.HtmlEncode(partLabel));
                b.Append("</b></div>");
            }
            b.Append(encoded);
            b.Append("</div></body></html>");
            return b.ToString();
        }

        private void ShowInfo(string message)
        {
            MessageBox.Show(this, message, "Outlook Part Sender", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(this, message, "Outlook Part Sender", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void ReleaseCom(object obj)
        {
            if (obj == null) return;
            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.FinalReleaseComObject(obj);
            }
            catch
            {
            }
        }
    }


    internal static class PortableProgram
    {
        [STAThread]
        private static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            object outlook = null;
            try
            {
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    MessageBox.Show(
                        "O Outlook Clássico não foi encontrado neste computador.\r\n\r\n" +
                        "Esta versão funciona somente com o Outlook Clássico para Windows.",
                        "Outlook Part Sender",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                try
                {
                    outlook = Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    outlook = Activator.CreateInstance(outlookType);
                }

                if (outlook == null)
                    throw new InvalidOperationException("Não foi possível abrir/conectar ao Outlook Clássico.");

                using (PartSenderForm form = new PartSenderForm(outlook, delegate { return false; }))
                {
                    System.Windows.Forms.Application.Run(form);
                }
            }
            catch (Exception ex)
            {
                try { Connect.Log("Portable Main ERROR: " + ex.ToString()); } catch { }
                MessageBox.Show(
                    "Não foi possível iniciar o Outlook Part Sender.\r\n\r\n" + ex.Message,
                    "Outlook Part Sender",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (outlook != null)
                {
                    try
                    {
                        if (Marshal.IsComObject(outlook))
                            Marshal.FinalReleaseComObject(outlook);
                    }
                    catch { }
                }
            }
        }
    }

}
