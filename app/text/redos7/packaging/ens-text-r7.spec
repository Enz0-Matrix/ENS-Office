# ENS-Tekst для РЕД ОС 7 («Муром»): окно Chromium в режиме приложения +
# помощник ens-text-r7 (файлы, сохранение на место, восстановление).
# Собирается в CI: rpmbuild -bb --define "_ensroot <каталог с файлами>" …

Name:           ens-text-r7
Version:        %{ens_version}
Release:        1
Summary:        ENS-Tekst — текстовый редактор ENS-Office для РЕД ОС 7
License:        Proprietary
URL:            https://github.com/Enz0-Matrix/ENS-Office
BuildArch:      x86_64
AutoReqProv:    no

%description
ENS-Tekst — текстовый редактор ENS-Office, совместимый с Microsoft Word (docx).
Версия для РЕД ОС 7: открывается своим окном (на основе установленного
Chromium), сохраняет файлы на место и восстанавливает документы после сбоя.
Работает без интернета.

%install
mkdir -p %{buildroot}/usr/bin %{buildroot}/usr/share/ens-text-r7 %{buildroot}/usr/share/applications %{buildroot}/usr/share/icons/hicolor/128x128/apps
install -m 0755 %{_ensroot}/ens-text-r7 %{buildroot}/usr/bin/ens-text-r7
cp -r %{_ensroot}/ui %{buildroot}/usr/share/ens-text-r7/ui
install -m 0644 %{_ensroot}/ens-text-r7.desktop %{buildroot}/usr/share/applications/ens-text-r7.desktop
install -m 0644 %{_ensroot}/ens-text-r7.png %{buildroot}/usr/share/icons/hicolor/128x128/apps/ens-text-r7.png

%files
/usr/bin/ens-text-r7
/usr/share/ens-text-r7
/usr/share/applications/ens-text-r7.desktop
/usr/share/icons/hicolor/128x128/apps/ens-text-r7.png

%post
update-desktop-database -q /usr/share/applications >/dev/null 2>&1 || :
gtk-update-icon-cache -q /usr/share/icons/hicolor >/dev/null 2>&1 || :

%postun
update-desktop-database -q /usr/share/applications >/dev/null 2>&1 || :
