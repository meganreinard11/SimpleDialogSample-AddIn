(function () {
    "use strict";
    var app = window.app || {};
    app.notification = {};
    $(document).ready(function () {    
        
         app.notification.MessageBanner = function (container) {          
             this.container = container;
             this.init();
        };
        
         app.notification.MessageBanner.prototype = CreatePrototype();
        
         $('.ms-MessageBanner').each(function () {
             new app.notification.MessageBanner(this);
         });

        function CreatePrototype(){
            var _clipper;
            var _bufferSize;
            var _textContainerMaxWidth = 700;
            var _clientWidth;
            var _textWidth;
            var _initTextWidth;
            var _chevronButton;
            var _errorBanner;
            var _closeButton;
            var _bufferElementsWidth = 88;
            var _bufferElementsWidthSmall = 35;
            var SMALL_BREAK_POINT = 480;
            
            var _onResize = function() {
                _clientWidth = _errorBanner.offsetWidth;
                if(window.innerWidth >= SMALL_BREAK_POINT ) {
                    _resizeRegular();
                } else {
                    _resizeSmall();
                }
            };
            
            var _resizeRegular = function() {
                if ((_clientWidth - _bufferSize) > _initTextWidth && _initTextWidth < _textContainerMaxWidth) {
                    _textWidth = "auto";
                     if (_chevronButton) { _chevronButton.className = "ms-MessageBanner-expand"; }
                    _collapse();
                } else {
                    _textWidth = Math.min((_clientWidth - _bufferSize), _textContainerMaxWidth) + "px";
                    if (_chevronButton) { if(_chevronButton.className.indexOf("is-visible") === -1) { _chevronButton.className += " is-visible"; } }
                }
                _clipper.style.width = _textWidth;
            };

            var _resizeSmall = function() {
                if (_clientWidth - (_bufferElementsWidthSmall + _closeButton.offsetWidth) > _initTextWidth) {
                    _textWidth = "auto";
                    _collapse();
                } else {
                    _textWidth = (_clientWidth - (_bufferElementsWidthSmall + _closeButton.offsetWidth)) + "px";
                }
                _clipper.style.width = _textWidth;
            };
          
            var _cacheDOM = function(context) {
                _errorBanner = context.container;
                _clipper = _errorBanner.querySelector('.ms-MessageBanner-clipper');
                _chevronButton = _errorBanner.querySelector('.ms-MessageBanner-expand');
                _bufferSize = _bufferElementsWidth;
                _closeButton = _errorBanner.querySelector('.ms-MessageBanner-close');
            };

            var _expand = function() {
                if (_chevronButton) {
                    var icon = _chevronButton.querySelector('.ms-Icon');
                    _errorBanner.className += " is-expanded";
                    icon.className = "ms-Icon ms-Icon--chevronsUp";
                } else {
                    _errorBanner.className += " is-expanded";
                }
            };

            var _collapse = function() {
                if (_chevronButton) {
                    var icon = _chevronButton.querySelector('.ms-Icon');
                    _errorBanner.className = "ms-MessageBanner";
                    icon.className = "ms-Icon ms-Icon--chevronsDown";
                } else {
                   _errorBanner.className = "ms-MessageBanner"; 
                }
            };

            var _toggleExpansion = function() {
                if (_errorBanner.className.indexOf("is-expanded") > -1) {
                    _collapse();
                } else {
                    _expand();
                }
            };

            var _hideBanner = function() {
                /* if(_errorBanner.className.indexOf("hide") === -1) {
                    _errorBanner.className += " hide";
                    setTimeout(function() { _errorBanner.className = "ms-MessageBanner is-hidden"; }, 500);
                } */
                if(_errorBanner.className.indexOf("hidden") === -1) {
                    _errorBanner.className += " hidden";
                }
            };

            var _showBanner = function() {
                _errorBanner.className = "ms-MessageBanner";
            };

           
            var _setListeners = function() {
                window.addEventListener('resize', _onResize, false);
                if (_chevronButton) { _chevronButton.addEventListener("click", _toggleExpansion, false); }
                _closeButton.addEventListener("click", _hideBanner, false);
            };

            var init = function() {
                _cacheDOM(this);
                _setListeners();
                _clientWidth = _errorBanner.offsetWidth;
                _initTextWidth = _clipper.offsetWidth;
                _onResize(null);
            };

            return {
                init: init,
                showBanner: _showBanner,
                hideBanner: _hideBanner,
                toggleExpansion: _toggleExpansion
            };
        };
    }); 
    window.app = app;
})();
