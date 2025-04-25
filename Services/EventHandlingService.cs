using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service class for managing PowerPoint application events
    /// </summary>
    public class EventHandlingService
    {
        private readonly PowerPoint.Application _application;
        private readonly Action<string, bool> _notificationCallback;
        private readonly TextFormattingService _textFormattingService;
        private readonly Action _refreshRibbonUICallback;
        private readonly ComObjectManager _comObjectManager;

        /// <summary>
        /// Initializes a new instance of the EventHandlingService class
        /// </summary>
        /// <param name="application">PowerPoint application instance</param>
        /// <param name="notificationCallback">Callback for displaying notifications</param>
        /// <param name="textFormattingService">Text formatting service for color cache resets</param>
        /// <param name="refreshRibbonUICallback">Callback to refresh ribbon UI</param>
        /// <param name="comObjectManager">COM object manager for handling COM object releases</param>
        public EventHandlingService(
            PowerPoint.Application application,
            Action<string, bool> notificationCallback,
            TextFormattingService textFormattingService,
            Action refreshRibbonUICallback,
            ComObjectManager comObjectManager)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notificationCallback = notificationCallback ?? throw new ArgumentNullException(nameof(notificationCallback));
            _textFormattingService = textFormattingService ?? throw new ArgumentNullException(nameof(textFormattingService));
            _refreshRibbonUICallback = refreshRibbonUICallback ?? throw new ArgumentNullException(nameof(refreshRibbonUICallback));
            _comObjectManager = comObjectManager ?? throw new ArgumentNullException(nameof(comObjectManager));
        }

        /// <summary>
        /// Subscribe to all required PowerPoint application events
        /// </summary>
        public void SubscribeToEvents()
        {
            try
            {
                // Theme and presentation events
                _application.ColorSchemeChanged += Application_ColorSchemeChanged;
                _application.PresentationSync += Application_PresentationSync;
                _application.WindowActivate += Application_WindowActivate;
                _application.PresentationOpen += Application_PresentationOpen;
                _application.AfterPresentationOpen += Application_AfterPresentationOpen;
                _application.WindowSelectionChange += Application_WindowSelectionChange;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error subscribing to events: {ex.Message}", true);
            }
        }

        /// <summary>
        /// Unsubscribe from all PowerPoint application events
        /// </summary>
        public void UnsubscribeFromEvents()
        {
            try
            {
                // Theme and presentation events
                _application.ColorSchemeChanged -= Application_ColorSchemeChanged;
                _application.PresentationSync -= Application_PresentationSync;
                _application.WindowActivate -= Application_WindowActivate;
                _application.PresentationOpen -= Application_PresentationOpen;
                _application.AfterPresentationOpen -= Application_AfterPresentationOpen;
                _application.WindowSelectionChange -= Application_WindowSelectionChange;
            }
            catch
            {
                // Ignore errors during unsubscription - we don't want to disrupt shutdown
            }
        }

        /// <summary>
        /// Event handler for color scheme changes
        /// </summary>
        private void Application_ColorSchemeChanged(PowerPoint.SlideRange SldRange)
        {
            try
            {
                // Reset color cache when theme colors change
                if (_textFormattingService != null)
                {
                    _textFormattingService.ResetColorCache();
                }

                // Refresh the ribbon UI to reflect theme changes
                _refreshRibbonUICallback();
            }
            catch
            {
                // Ignore errors to avoid disrupting PowerPoint
            }
            finally
            {
                // Release the SlideRange COM object
                if (SldRange != null)
                {
                    _comObjectManager.ReleaseComObject(SldRange, "SlideRange from ColorSchemeChanged");
                }
            }
        }

        /// <summary>
        /// Event handler for presentation sync - occurs when a presentation is synchronized with a server
        /// </summary>
        private void Application_PresentationSync(PowerPoint.Presentation Pres, Office.MsoSyncEventType SyncEventType)
        {
            try
            {
                // Reset color cache as theme might have changed during sync
                if (_textFormattingService != null)
                {
                    _textFormattingService.ResetColorCache();
                }

                // Refresh the ribbon UI
                _refreshRibbonUICallback();
            }
            catch
            {
                // Ignore errors to avoid disrupting PowerPoint
            }
            finally
            {
                // Release the Presentation COM object
                if (Pres != null)
                {
                    _comObjectManager.ReleaseComObject(Pres, "Presentation from PresentationSync");
                }
            }
        }

        /// <summary>
        /// Event handler for window activation - occurs when a document window is activated
        /// </summary>
        private void Application_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            try
            {
                // Reset color cache as the active presentation/theme might have changed
                if (_textFormattingService != null)
                {
                    _textFormattingService.ResetColorCache();
                }

                // Refresh the ribbon UI to reflect potential theme changes
                _refreshRibbonUICallback();
            }
            catch
            {
                // Ignore errors to avoid disrupting PowerPoint
            }
            finally
            {
                // Release COM objects
                if (Wn != null)
                {
                    _comObjectManager.ReleaseComObject(Wn, "DocumentWindow from WindowActivate");
                }

                if (Pres != null)
                {
                    _comObjectManager.ReleaseComObject(Pres, "Presentation from WindowActivate");
                }
            }
        }

        /// <summary>
        /// Event handler for presentation open
        /// </summary>
        private void Application_PresentationOpen(PowerPoint.Presentation presentation)
        {
            try
            {
                // Reset the text formatting service's color cache
                if (_textFormattingService != null)
                {
                    _textFormattingService.ResetColorCache();
                }

                // Refresh the ribbon UI
                _refreshRibbonUICallback();
            }
            catch
            {
                // Ignore errors to avoid disrupting PowerPoint
            }
            finally
            {
                // Release the Presentation COM object
                if (presentation != null)
                {
                    _comObjectManager.ReleaseComObject(presentation, "Presentation from PresentationOpen");
                }
            }
        }

        /// <summary>
        /// Event handler for after presentation open
        /// </summary>
        private void Application_AfterPresentationOpen(PowerPoint.Presentation presentation)
        {
            try
            {
                // Reset the text formatting service's color cache
                if (_textFormattingService != null)
                {
                    _textFormattingService.ResetColorCache();
                }

                // Refresh the ribbon UI
                _refreshRibbonUICallback();
            }
            catch
            {
                // Ignore errors to avoid disrupting PowerPoint
            }
            finally
            {
                // Release the Presentation COM object
                if (presentation != null)
                {
                    _comObjectManager.ReleaseComObject(presentation, "Presentation from AfterPresentationOpen");
                }
            }
        }

        /// <summary>
        /// Event handler for window selection change
        /// May indicate theme changes as well
        /// </summary>
        private void Application_WindowSelectionChange(PowerPoint.Selection selection)
        {
            try
            {
                // Check if the active presentation window exists
                if (_application.ActiveWindow != null &&
                    _application.ActiveWindow.Presentation != null)
                {
                    // Refresh the ribbon UI to reflect potential theme changes
                    _refreshRibbonUICallback();
                }
            }
            catch
            {
                // Ignore errors to avoid disrupting PowerPoint
            }
            finally
            {
                // Release the Selection COM object
                if (selection != null)
                {
                    _comObjectManager.ReleaseComObject(selection, "Selection from WindowSelectionChange");
                }
            }
        }
    }
}