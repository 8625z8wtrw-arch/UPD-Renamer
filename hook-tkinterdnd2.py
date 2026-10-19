from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

datas = collect_data_files('tkinterdnd2')
bin_datas = collect_dynamic_libs('tkinterdnd2')
datas.extend(bin_datas)