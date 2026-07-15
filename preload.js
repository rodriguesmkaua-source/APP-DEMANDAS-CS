const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('csApp', {
  loadData:           ()       => ipcRenderer.invoke('load-data'),
  saveData:           (items)  => ipcRenderer.invoke('save-data', items),
  backup:             (items)  => ipcRenderer.invoke('backup', items),
  restore:            ()       => ipcRenderer.invoke('restore'),
  exportExcel:        (data)   => ipcRenderer.invoke('export-excel', data),
  openHistorico:      ()       => ipcRenderer.invoke('open-historico'),
  notify:             (opts)   => ipcRenderer.invoke('notify', opts),
  chooseExportFolder: ()       => ipcRenderer.invoke('choose-export-folder'),
  getExportFolder:    ()       => ipcRenderer.invoke('get-export-folder'),
  openFechamento:     (data)   => ipcRenderer.invoke('open-fechamento', data),
  getFechamentoData:  ()       => ipcRenderer.invoke('get-fechamento-data'),
  printFechamentoPDF: ()       => ipcRenderer.invoke('print-fechamento-pdf'),
  getOperatorLogo:    (op)     => ipcRenderer.invoke('get-operator-logo', op),
  captureFechamentoPNG:()      => ipcRenderer.invoke('capture-fechamento-png'),
  getCoverImage:      ()       => ipcRenderer.invoke('get-cover-image'),
  openFechamentoCompleto:     (data) => ipcRenderer.invoke('open-fechamento-completo', data),
  getFechamentoCompletoData:  ()     => ipcRenderer.invoke('get-fechamento-completo-data'),
  printFechamentoCompletoPDF: ()     => ipcRenderer.invoke('print-fechamento-completo-pdf'),
})
