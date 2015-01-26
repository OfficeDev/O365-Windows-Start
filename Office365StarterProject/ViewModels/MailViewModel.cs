// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Office365StarterProject.Common;
using Office365StarterProject.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Input;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;

namespace Office365StarterProject.ViewModels
{
    /// <summary>
    /// Defines the mail view model.
    /// </summary>
    class MailViewModel : ViewModelBase
    {
        private MailOperations _mailOperations = null;
        private MailItemViewModel _selectedMail = null;
        private string _newMailSubject = null;
        private string _newMailRecipients = null;
        private string _newMailBodyContent = null;
        private bool _loadingMail = false;
        private int _currentPage = 1;
        private int _previousPage = -1;
        private bool _isLastPage = false;
        private int _pageSize = 10;

        public MailViewModel()
        {
 
            // Instantiate a private instance of the mail operations object
            _mailOperations = new MailOperations();

            this.MailItems = new ObservableCollection<MailItemViewModel>();

            //construct relay commands to be bound to controls on a UI
            this.SendMailCommand = new RelayCommand(ExecuteSendMailCommandAsync);
            this.GetMailCommand = new RelayCommand(ExecuteGetMailCommandAsync);
            this.GetPrevPageCommand = new RelayCommand(ExecuteGetPrevPageCommandAsync, CanGetPrevPage);
            this.DeleteMailCommand = new RelayCommand(ExecuteDeleteMailCommandAsync, CanDeleteMail);
        }

        /// <summary>
        /// The mail items on a bound UI list
        /// </summary>
        public ObservableCollection<MailItemViewModel> MailItems { get; private set; }

        /// <summary>
        /// Command to send an email.
        /// </summary>
        public ICommand SendMailCommand { protected set; get; }

        /// <summary>
        /// Command to get the user's email.
        /// </summary>
        public ICommand GetMailCommand { protected set; get; }

        public ICommand GetPrevPageCommand { protected set; get; }

        /// <summary>
        /// Command to delete a mail item.
        /// </summary>
        public ICommand DeleteMailCommand { protected set; get; }

        /// <summary>
        /// Gets or sets whether we are in the process of loading mail.
        /// </summary>
        public bool LoadingMail
        {
            get
            {
                return _loadingMail;
            }
            private set
            {
                SetProperty(ref _loadingMail, value);
            }
        }

        public string NewMailSubject
        {
            get
            {
                return _newMailSubject;
            }
           set
            {
                SetProperty(ref _newMailSubject, value);
            }

        }

        public string NewMailRecipients
        {

            get
            {
                return _newMailRecipients;
            }
            set
            {
                SetProperty(ref _newMailRecipients, value);
            }
        }

        public string NewMailBodyContent
        {
            get
            {
                return _newMailBodyContent;
            }
            set
            {
                SetProperty(ref _newMailBodyContent, value);
            }

        }

        /// <summary>
        /// Sets or gets the selected MailItemViewModel from the mail list in a UI.
        /// Updates mail item view model fields bound to mail field properties exposed in this model.
        /// </summary>
        public MailItemViewModel SelectedMail
        {
            get
            {
                return _selectedMail;
            }
            set
            {
                if (SetProperty(ref _selectedMail, value))
                {

                    // Enable and disable the delete mail command depending on whether a mail message has been selected.
                    ((RelayCommand)this.DeleteMailCommand).RaiseCanExecuteChanged();
                }
            }
        }

        private bool CanDeleteMail()
        {
            return (this.SelectedMail != null);
        }

        private bool CanGetPrevPage()
        {
            return (_previousPage > 0);
        }

        /// <summary>
        /// Sends a mail item and adds it to the collection. 
        /// </summary>
        /// <remarks>The mail item is created locally.</remarks>
        async void ExecuteSendMailCommandAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_newMailRecipients))
                {
                    LoggingViewModel.Instance.Information = "Please include at least one recipient.";
                }
                else
                {
                    await _mailOperations.ComposeAndSendMailAsync(_newMailSubject, _newMailBodyContent, _newMailRecipients);
                    LoggingViewModel.Instance.Information = "Your mail was sent.";
                }
            }

            catch (Exception ex)
            {
                LoggingViewModel.Instance.Information = "Error sending mail: " + ex.Message;
            }
        }


        /// <summary>
        /// Gets the user's email from the Exchange service.
        /// </summary>
        async void ExecuteGetMailCommandAsync()
        {
            this.LoadingMail = true;

            //If user has clicked on this button after viewing the last page, reset the pagination values.
            if (_isLastPage)
            {
                _currentPage = 1;
                _previousPage = -1;
                _isLastPage = false;
            }
            await this.LoadMailAsync(_currentPage);
            this.LoadingMail = false;
            _currentPage++;
            _previousPage++;
 
            // Enable the previous page button if the value of _previousPage is above 0.
            ((RelayCommand)this.GetPrevPageCommand).RaiseCanExecuteChanged();
        }

        /// <summary>
        /// Gets the previous page of the user's email from the Exchange service.
        /// </summary>
        async void ExecuteGetPrevPageCommandAsync()
        {

            //If the user has clicked on this button after viewing the last page, make sure that
            //the value of _isLastPage is set back to false.
            if (_isLastPage)
            {
                _isLastPage = false;
            }
            this.LoadingMail = true;
            await this.LoadMailAsync(_previousPage);
            this.LoadingMail = false;
            _previousPage--;
            _currentPage--;

            // Disable the previous page button if the value of _previousPage drops below 1.
            ((RelayCommand)this.GetPrevPageCommand).RaiseCanExecuteChanged();
        }


        private async Task<bool> LoadMailAsync(int page)
        {
            LoggingViewModel.Instance.Information = string.Empty;

            try
            {
 
                //Clear out any mail added in previous calls to LoadMailAsync()
                if (MailItems != null)
                    MailItems.Clear();
                else
                    MailItems = new ObservableCollection<MailItemViewModel>();

                LoggingViewModel.Instance.Information = "Getting mail ...";


                //Get mail from Exchange service via API.
                var mail = await _mailOperations.GetEmailMessagesAsync(page, _pageSize);

                if (mail.Count == 0 && _currentPage == 1)
                {
                    LoggingViewModel.Instance.Information = "You have no mail.";
                    _isLastPage = true;
                }
                else if (mail.Count == 0)
                {
                    LoggingViewModel.Instance.Information = "You have no more mail. Click the \"Get Items\" button to reload the first page.";
                    _isLastPage = true;
                }
                else
                {
                    // Load emails into the observable collection that is bound to UI
                    foreach (var mailItem in mail)
                    {
                        MailItems.Add(new MailItemViewModel(mailItem));
                    }

                    if (mail.Count < _pageSize)
                    {
                        LoggingViewModel.Instance.Information = String.Format("{0} mail items loaded. Click the \"Get Items\" button to reload the first page.", MailItems.Count);
                        _isLastPage = true;
                    }
                    else
                    {
                        LoggingViewModel.Instance.Information = String.Format("{0} mail items loaded. Click the \"Get Items\" button for more.", MailItems.Count);

                    }

                }
            }
            catch (Exception ex)
            {
                LoggingViewModel.Instance.Information = "Error loading mail: " + ex.Message;
                return false;
            }
            return true;
        }


        /// <summary>
        /// Sends mail item remove request to the Exchange service.
        /// </summary>
        async void ExecuteDeleteMailCommandAsync()
        {
            try
            {
                if (await MessageDialogHelper.ShowYesNoDialogAsync(String.Format("Are you sure you want to delete the mail item '{0}'?", this._selectedMail.Subject), "Confirm Deletion"))
                {
                    if (!String.IsNullOrEmpty(this._selectedMail.ID))
                    {
                        if (await _mailOperations.DeleteMailItemAsync(this._selectedMail.ID))

                            //Removes email from bound observable collection
                            MailItems.Remove((MailItemViewModel)_selectedMail);

                    }

                }
            }
            catch (Exception)
            {
                LoggingViewModel.Instance.Information = "We could not delete your mail item.";
            }
        }

    }
}
//********************************************************* 
// 
//O365-APIs-Start-Windows, https://github.com/OfficeDev/O365-APIs-Start-Windows
//
//Copyright (c) Microsoft Corporation
//All rights reserved. 
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 
