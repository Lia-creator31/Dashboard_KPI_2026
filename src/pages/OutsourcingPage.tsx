// import React, { useState, useEffect, useCallback, useMemo } from 'react';
// import { departmentsData, Department } from '../data';
// import { supabase } from '../lib/supabase';
// import { UserSession } from '../App';
// import {
//   Building2, HardHat, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
//   Pencil, Trash2, X, Check, Clock, Plus, Calendar, Briefcase
// } from 'lucide-react';

// interface TaskItem {
//   id: string;
//   biroName: string;
//   project: string;
//   taskName: string;
//   startDate: string;
//   endDate: string;
//   pic: string;
//   jo: string;
//   kodeJc: string;
//   rev?: string;
//   release?: string;
// }

// interface TimesheetLine {
//   id?: string;
//   date: string;
//   workOrderId: string;
//   workOrderCode: string;
//   description: string;
//   effectiveHours: string;
//   overtimeHours: string;
//   plannedHour?: string;
//   idleHours?: string;
//   type?: string;
//   project?: string;
//   wbsActivity?: string;
//   analyticAccount?: string;
//   day?: string;
//   isLembur?: boolean;
// }

// interface Timesheet {
//   id: string;
//   kodeTimesheet: string;
//   userId: string;
//   userName: string;
//   dateStart: string;
//   dateEnd: string;
//   status: 'draft' | 'submitted' | 'approved';
//   lines: TimesheetLine[];
//   createdAt: string;
// }

// interface OutsourcingPageProps {
//   user: UserSession;
//   onLogout: () => void;
// }

// function generateTimesheetCode(): string {
//   const timestamp = Date.now();
//   const random = Math.floor(Math.random() * 1000);
//   return `TIM-${timestamp}${random}`;
// }

// function formatDateToDay(dateStr: string): string {
//   const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
//   const date = new Date(dateStr);
//   return days[date.getDay()];
// }

// export default function OutsourcingPage({ user, onLogout }: OutsourcingPageProps) {
//   const [selectedDept, setSelectedDept] = useState<Department | null>(null);
//   const [selectedBiro, setSelectedBiro] = useState<string | null>(null);
//   const [searchQuery, setSearchQuery] = useState('');
  
//   // State untuk Timesheet
//   const [timesheets, setTimesheets] = useState<Timesheet[]>([]);
//   const [isCreatingTimesheet, setIsCreatingTimesheet] = useState(false);
//   const [currentTimesheet, setCurrentTimesheet] = useState<Partial<Timesheet>>({
//     dateStart: '',
//     dateEnd: '',
//     lines: [],
//   });
//   const [isAddingLine, setIsAddingLine] = useState(false);
//   const [currentLine, setCurrentLine] = useState<Partial<TimesheetLine>>({
//     date: '',
//     effectiveHours: '',
//     overtimeHours: '',
//     plannedHour: '00:00',
//     idleHours: '00:00',
//   });
//   const [availableWorkOrders, setAvailableWorkOrders] = useState<TaskItem[]>([]);
//   const [selectedLineIndex, setSelectedLineIndex] = useState<number | null>(null);

//   // Load Timesheet dari Supabase
//   const loadTimesheets = useCallback(async () => {
//     try {
//       const { data, error } = await supabase
//         .from('outsourcing_timesheets')
//         .select('*')
//         .eq('user_id', user.nip)
//         .order('created_at', { ascending: false });
      
//       if (error) throw error;
      
//       // Load lines untuk setiap timesheet
//       const timesheetsWithLines = await Promise.all(
//         (data || []).map(async (ts) => {
//           const { data: lines } = await supabase
//             .from('outsourcing_timesheet_lines')
//             .select('*')
//             .eq('timesheet_id', ts.id);
          
//           return {
//             ...ts,
//             lines: lines || [],
//           };
//         })
//       );
      
//       setTimesheets(timesheetsWithLines as Timesheet[]);
//     } catch (error) {
//       console.error('Error loading timesheets:', error);
//     }
//   }, [user.nip]);

// const loadAvailableWorkOrders = useCallback(async (startDate: string, endDate: string) => {
//   try {
//     console.log('🔍 Mencari work order untuk:', user.nama, '| Tanggal:', startDate);
    
//     // 1. Ambil SEMUA work order yang statusnya 'approved' dulu
//     const { data, error } = await supabase
//       .from('job_cards')
//       .select('*')
//       .eq('status', 'approved');
    
//     if (error) {
//       console.error('❌ Error dari Supabase:', error);
//       throw error;
//     }
    
//     console.log('✅ Total work order approved ditemukan:', data?.length || 0, data);

//     // 2. Filter di client-side (Browser) agar lebih fleksibel
//     const filteredData = (data || []).filter(wo => {
//       // A. Cek kecocokan Nama PIC (handle jika ada NIP atau hanya nama)
//       const userNameLower = user.nama.toLowerCase();
//       const userNipLower = user.nip.toLowerCase();
//       const picLower = wo.pic.toLowerCase();
      
//       const isNameMatch = picLower.includes(userNameLower) || 
//                           picLower.includes(userNipLower) || 
//                           userNameLower.includes(picLower);

//       // B. Cek kecocokan Tanggal (BANDINGKAN SEBAGAI STRING YYYY-MM-DD)
//       // Ini mencegah bug timezone dari new Date()
//       const isDateMatch = wo.start_date <= startDate && wo.end_date >= startDate;
      
//       return isNameMatch && isDateMatch;
//     });
    
//     console.log('🎯 Work order yang lolos filter (seharusnya DP5 ada di sini):', filteredData);
//     setAvailableWorkOrders(filteredData);
    
//   } catch (error) {
//     console.error('❌ Gagal memuat work order:', error);
//   }
// }, [user.nama, user.nip]);

//   useEffect(() => {
//     loadTimesheets();
//   }, [loadTimesheets]);

//   // Handler untuk membuat timesheet baru
//   const handleCreateTimesheet = () => {
//     setIsCreatingTimesheet(true);
//     setCurrentTimesheet({
//       dateStart: '',
//       dateEnd: '',
//       lines: [],
//     });
//   };

//   // Handler untuk menyimpan timesheet
//   const handleSaveTimesheet = async () => {
//     if (!currentTimesheet.dateStart || !currentTimesheet.dateEnd) {
//       alert('Tanggal mulai dan tanggal akhir harus diisi!');
//       return;
//     }

//     try {
//       const newTimesheet = {
//         kode_timesheet: generateTimesheetCode(),
//         user_id: user.nip,
//         user_name: user.nama,
//         date_start: currentTimesheet.dateStart,
//         date_end: currentTimesheet.dateEnd,
//         status: 'draft',
//       };

//       const { data, error } = await supabase
//         .from('outsourcing_timesheets')
//         .insert(newTimesheet)
//         .select()
//         .single();

//       if (error) throw error;

//       // Simpan lines
//       if (currentTimesheet.lines && currentTimesheet.lines.length > 0) {
//         const linesToInsert = currentTimesheet.lines.map(line => ({
//           timesheet_id: data.id,
//           date: line.date,
//           work_order_id: line.workOrderId,
//           work_order_code: line.workOrderCode,
//           description: line.description,
//           effective_hours: line.effectiveHours,
//           overtime_hours: line.overtimeHours,
//           planned_hour: line.plannedHour || '00:00',
//           idle_hours: line.idleHours || '00:00',
//         }));

//         await supabase.from('outsourcing_timesheet_lines').insert(linesToInsert);
//       }

//       alert('Timesheet berhasil disimpan!');
//       setIsCreatingTimesheet(false);
//       loadTimesheets();
//     } catch (error) {
//       console.error('Error saving timesheet:', error);
//       alert('Gagal menyimpan timesheet');
//     }
//   };

//   // Handler untuk add line
//   const handleAddLine = () => {
//     if (!currentTimesheet.dateStart || !currentTimesheet.dateEnd) {
//       alert('Silakan isi tanggal mulai dan tanggal akhir terlebih dahulu!');
//       return;
//     }
    
//     setIsAddingLine(true);
//     loadAvailableWorkOrders(currentTimesheet.dateStart!, currentTimesheet.dateEnd!);
//     setCurrentLine({
//       date: '',
//       effectiveHours: '',
//       overtimeHours: '',
//       plannedHour: '00:00',
//       idleHours: '00:00',
//     });
//   };

//   // Handler untuk menyimpan line
//   const handleSaveLine = () => {
//     if (!currentLine.date || !currentLine.workOrderId) {
//       alert('Tanggal dan Work Order harus diisi!');
//       return;
//     }

//     const workOrder = availableWorkOrders.find(wo => wo.id === currentLine.workOrderId);
    
//     const newLine: TimesheetLine = {
//       date: currentLine.date!,
//       workOrderId: currentLine.workOrderId!,
//       workOrderCode: workOrder?.kodeJc || '',
//       description: currentLine.description || '',
//       effectiveHours: currentLine.effectiveHours || '0',
//       overtimeHours: currentLine.overtimeHours || '0',
//       plannedHour: currentLine.plannedHour || '00:00',
//       idleHours: currentLine.idleHours || '00:00',
//       day: formatDateToDay(currentLine.date),
//     };

//     if (selectedLineIndex !== null) {
//       // Update existing line
//       const updatedLines = [...(currentTimesheet.lines || [])];
//       updatedLines[selectedLineIndex] = newLine;
//       setCurrentTimesheet({ ...currentTimesheet, lines: updatedLines });
//     } else {
//       // Add new line
//       setCurrentTimesheet({
//         ...currentTimesheet,
//         lines: [...(currentTimesheet.lines || []), newLine],
//       });
//     }

//     setIsAddingLine(false);
//     setSelectedLineIndex(null);
//     setCurrentLine({
//       date: '',
//       effectiveHours: '',
//       overtimeHours: '',
//       plannedHour: '00:00',
//       idleHours: '00:00',
//     });
//   };

//   // Handler untuk edit line
//   const handleEditLine = (index: number) => {
//     const line = currentTimesheet.lines?.[index];
//     if (line) {
//       setCurrentLine(line);
//       setSelectedLineIndex(index);
//       setIsAddingLine(true);
//       loadAvailableWorkOrders(currentTimesheet.dateStart!, currentTimesheet.dateEnd!);
//     }
//   };

//   // Handler untuk delete line
//   const handleDeleteLine = (index: number) => {
//     if (confirm('Hapus baris ini?')) {
//       const updatedLines = currentTimesheet.lines?.filter((_, i) => i !== index) || [];
//       setCurrentTimesheet({ ...currentTimesheet, lines: updatedLines });
//     }
//   };

//   // Handler untuk submit timesheet
//   const handleSubmitTimesheet = async (timesheetId: string) => {
//     try {
//       await supabase
//         .from('outsourcing_timesheets')
//         .update({ status: 'submitted' })
//         .eq('id', timesheetId);
      
//       alert('Timesheet berhasil disubmit!');
//       loadTimesheets();
//     } catch (error) {
//       console.error('Error submitting timesheet:', error);
//       alert('Gagal submit timesheet');
//     }
//   };

//   // Filter work order berdasarkan tanggal
//   const getWorkOrdersForDate = (date: string) => {
//     return availableWorkOrders.filter(wo => {
//       return date >= wo.startDate && date <= wo.endDate;
//     });
//   };

//   // Render halaman utama outsourcing
//   if (!selectedDept) {
//     return (
//       <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
//         <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
//           <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
//             <div className="flex items-center gap-2.5">
//               <div className="p-1.5 bg-amber-600 rounded-lg text-white">
//                 <HardHat className="w-4 h-4" />
//               </div>
//               <span className="font-bold text-sm text-white">OUTSOURCING TIMESHEET</span>
//             </div>
//             <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
//               Keluar
//             </button>
//           </div>
//         </header>

//         <main className="max-w-6xl mx-auto px-4 py-8">
//           <div className="flex items-center justify-between mb-6">
//             <div>
//               <h1 className="text-2xl font-bold text-white">Project Timesheet</h1>
//               <p className="text-slate-400 text-sm">Kelola timesheet kerja Anda</p>
//             </div>
//             <button
//               onClick={handleCreateTimesheet}
//               className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-sm font-semibold flex items-center gap-2 cursor-pointer"
//             >
//               <Plus className="w-4 h-4" /> Create
//             </button>
//           </div>

//           {/* Daftar Timesheet */}
//           <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
//             {timesheets.map((ts) => (
//               <div key={ts.id} className="bg-slate-900 border border-slate-800 rounded-xl p-4 hover:border-blue-500 transition">
//                 <div className="flex items-start justify-between mb-3">
//                   <h3 className="font-mono font-bold text-blue-400">{ts.kodeTimesheet}</h3>
//                   <span className={`px-2 py-1 rounded text-xs font-semibold ${
//                     ts.status === 'approved' ? 'bg-emerald-500/10 text-emerald-400' :
//                     ts.status === 'submitted' ? 'bg-amber-500/10 text-amber-400' :
//                     'bg-slate-700 text-slate-300'
//                   }`}>
//                     {ts.status === 'approved' ? 'Approved' : ts.status === 'submitted' ? 'Submitted' : 'Draft'}
//                   </span>
//                 </div>
                
//                 <div className="space-y-2 text-xs">
//                   <div>
//                     <span className="text-slate-400">Responsible:</span>
//                     <p className="text-white font-medium">{ts.userName}</p>
//                   </div>
//                   <div className="flex justify-between">
//                     <span className="text-slate-400">Start Date</span>
//                     <span className="text-emerald-400">{ts.dateStart}</span>
//                   </div>
//                   <div className="flex justify-between">
//                     <span className="text-slate-400">End Date</span>
//                     <span className="text-emerald-400">{ts.dateEnd}</span>
//                   </div>
//                   <div className="flex justify-between pt-2 border-t border-slate-800">
//                     <span className="text-slate-400">Total Lines</span>
//                     <span className="text-white">{ts.lines?.length || 0}</span>
//                   </div>
//                 </div>

//                 {ts.status === 'draft' && (
//                   <div className="mt-4 flex gap-2">
//                     <button
//                       onClick={() => handleSubmitTimesheet(ts.id)}
//                       className="flex-1 px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs cursor-pointer"
//                     >
//                       Submit
//                     </button>
//                   </div>
//                 )}
//               </div>
//             ))}
//           </div>
//         </main>

//         {/* Modal Create Timesheet */}
//         {isCreatingTimesheet && (
//           <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
//             <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-4xl max-h-[90vh] overflow-y-auto">
//               <div className="p-4 border-b border-slate-800 flex items-center justify-between">
//                 <h2 className="text-lg font-bold text-white">New</h2>
//                 <button onClick={() => setIsCreatingTimesheet(false)} className="text-slate-400 hover:text-white cursor-pointer">
//                   <X className="w-5 h-5" />
//                 </button>
//               </div>

//               <div className="p-6 space-y-6">
//                 {/* Header Info */}
//                 <div className="grid grid-cols-2 gap-4 text-sm">
//                   <div>
//                     <label className="block text-slate-400 mb-1">User</label>
//                     <input
//                       type="text"
//                       value={`${user.nip} - ${user.nama}`}
//                       readOnly
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400"
//                     />
//                   </div>
//                   <div>
//                     <label className="block text-slate-400 mb-1">Unit Kerja</label>
//                     <input
//                       type="text"
//                       value="Outsourcing"
//                       readOnly
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400"
//                     />
//                   </div>
//                   <div>
//                     <label className="block text-slate-400 mb-1">Date Start</label>
//                     <input
//                       type="date"
//                       value={currentTimesheet.dateStart || ''}
//                       onChange={(e) => setCurrentTimesheet({ ...currentTimesheet, dateStart: e.target.value })}
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
//                     />
//                   </div>
//                   <div>
//                     <label className="block text-slate-400 mb-1">Date End</label>
//                     <input
//                       type="date"
//                       value={currentTimesheet.dateEnd || ''}
//                       onChange={(e) => setCurrentTimesheet({ ...currentTimesheet, dateEnd: e.target.value })}
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
//                     />
//                   </div>
//                 </div>

//                 {/* Project Timesheet Line */}
//                 <div>
//                   <div className="flex items-center justify-between mb-3">
//                     <h3 className="text-sm font-semibold text-slate-300">Project Timesheet Line</h3>
//                     <button
//                       onClick={handleAddLine}
//                       className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs cursor-pointer flex items-center gap-1"
//                     >
//                       <Plus className="w-3 h-3" /> Add a line
//                     </button>
//                   </div>

//                   <div className="bg-slate-950 border border-slate-800 rounded-lg overflow-hidden">
//                     <table className="w-full text-xs">
//                       <thead className="bg-slate-900">
//                         <tr>
//                           <th className="px-3 py-2 text-left text-slate-400">Date</th>
//                           <th className="px-3 py-2 text-left text-slate-400">Work Order</th>
//                           <th className="px-3 py-2 text-left text-slate-400">Description</th>
//                           <th className="px-3 py-2 text-right text-slate-400">Effective Hours</th>
//                           <th className="px-3 py-2 text-right text-slate-400">Overtime Hours</th>
//                           <th className="px-3 py-2 text-center text-slate-400">Aksi</th>
//                         </tr>
//                       </thead>
//                       <tbody className="divide-y divide-slate-800">
//                         {currentTimesheet.lines?.map((line, index) => (
//                           <tr key={index} className="hover:bg-slate-900/50">
//                             <td className="px-3 py-2 text-white">{line.date}</td>
//                             <td className="px-3 py-2 text-emerald-400 font-mono">{line.workOrderCode}</td>
//                             <td className="px-3 py-2 text-slate-300">{line.description}</td>
//                             <td className="px-3 py-2 text-right text-white">{line.effectiveHours}</td>
//                             <td className="px-3 py-2 text-right text-amber-400">{line.overtimeHours}</td>
//                             <td className="px-3 py-2 text-center">
//                               <div className="flex justify-center gap-1">
//                                 <button
//                                   onClick={() => handleEditLine(index)}
//                                   className="p-1 text-blue-400 hover:bg-blue-500/20 rounded cursor-pointer"
//                                 >
//                                   <Pencil className="w-3 h-3" />
//                                 </button>
//                                 <button
//                                   onClick={() => handleDeleteLine(index)}
//                                   className="p-1 text-rose-400 hover:bg-rose-500/20 rounded cursor-pointer"
//                                 >
//                                   <Trash2 className="w-3 h-3" />
//                                 </button>
//                               </div>
//                             </td>
//                           </tr>
//                         ))}
//                         {(!currentTimesheet.lines || currentTimesheet.lines.length === 0) && (
//                           <tr>
//                             <td colSpan={6} className="px-3 py-8 text-center text-slate-500 text-xs">
//                               Click "Add a line" to add timesheet entries
//                             </td>
//                           </tr>
//                         )}
//                       </tbody>
//                     </table>
//                   </div>
//                 </div>

//                 {/* Total Hours */}
//                 <div className="flex justify-end">
//                   <div className="text-sm">
//                     <span className="text-slate-400">Total Hours: </span>
//                     <span className="font-bold text-white">
//                       {currentTimesheet.lines?.reduce((acc, line) => {
//                         const eff = parseFloat(line.effectiveHours) || 0;
//                         const ot = parseFloat(line.overtimeHours) || 0;
//                         return acc + eff + ot;
//                       }, 0).toFixed(2) || '0.00'}
//                     </span>
//                   </div>
//                 </div>

//                 {/* Action Buttons */}
//                 <div className="flex justify-end gap-3 pt-4 border-t border-slate-800">
//                   <button
//                     onClick={() => setIsCreatingTimesheet(false)}
//                     className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg cursor-pointer"
//                   >
//                     Discard
//                   </button>
//                   <button
//                     onClick={handleSaveTimesheet}
//                     className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg cursor-pointer"
//                   >
//                     Save
//                   </button>
//                 </div>
//               </div>
//             </div>
//           </div>
//         )}

//         {/* Modal Add/Edit Line */}
//         {isAddingLine && (
//           <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
//             <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
//               <div className="p-4 border-b border-slate-800 flex items-center justify-between">
//                 <h2 className="text-lg font-bold text-white">Create Lines</h2>
//                 <button onClick={() => { setIsAddingLine(false); setSelectedLineIndex(null); }} className="text-slate-400 hover:text-white cursor-pointer">
//                   <X className="w-5 h-5" />
//                 </button>
//               </div>

//               <div className="p-6 space-y-4">
//                 <div className="grid grid-cols-2 gap-4">
//                   <div className="col-span-2">
//                     <label className="block text-slate-400 mb-1 text-sm">Type</label>
//                     <select className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm">
//                       <option value="">Select Type</option>
//                       <option value="regular">Regular</option>
//                       <option value="overtime">Overtime</option>
//                     </select>
//                   </div>

//                   <div className="col-span-2">
//                     <label className="block text-slate-400 mb-1 text-sm">Work Center</label>
//                     <input
//                       type="text"
//                       value="Outsourcing"
//                       readOnly
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400 text-sm"
//                     />
//                   </div>

//                   <div className="col-span-2">
//                     <label className="block text-slate-400 mb-1 text-sm">User</label>
//                     <input
//                       type="text"
//                       value={`${user.nip} - ${user.nama}`}
//                       readOnly
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400 text-sm"
//                     />
//                   </div>

//                   <div className="col-span-2">
//                     <label className="block text-slate-400 mb-1 text-sm">Date</label>
//                     <input
//                       type="date"
//                       value={currentLine.date || ''}
//                       onChange={(e) => setCurrentLine({ ...currentLine, date: e.target.value })}
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
//                     />
//                   </div>

//                   <div className="col-span-2">
//                     <label className="block text-slate-400 mb-1 text-sm">Work Order</label>
//                     <select
//                       value={currentLine.workOrderId || ''}
//                       onChange={(e) => {
//                         const wo = availableWorkOrders.find(w => w.id === e.target.value);
//                         setCurrentLine({
//                           ...currentLine,
//                           workOrderId: e.target.value,
//                           description: wo?.taskName || '',
//                         });
//                       }}
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
//                     >
//                       <option value="">Select Work Order</option>
//                       {availableWorkOrders.map((wo) => (
//                         <option key={wo.id} value={wo.id}>
//                           {wo.kodeJc} - {wo.taskName}
//                         </option>
//                       ))}
//                     </select>
//                   </div>

//                   <div className="col-span-2">
//                     <label className="block text-slate-400 mb-1 text-sm">Description</label>
//                     <textarea
//                       value={currentLine.description || ''}
//                       onChange={(e) => setCurrentLine({ ...currentLine, description: e.target.value })}
//                       rows={3}
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
//                     />
//                   </div>

//                   <div>
//                     <label className="block text-slate-400 mb-1 text-sm">Effective Hours</label>
//                     <input
//                       type="text"
//                       value={currentLine.effectiveHours || ''}
//                       onChange={(e) => setCurrentLine({ ...currentLine, effectiveHours: e.target.value })}
//                       placeholder="00:00"
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
//                     />
//                   </div>

//                   <div>
//                     <label className="block text-slate-400 mb-1 text-sm">Overtime Hours</label>
//                     <input
//                       type="text"
//                       value={currentLine.overtimeHours || ''}
//                       onChange={(e) => setCurrentLine({ ...currentLine, overtimeHours: e.target.value })}
//                       placeholder="00:00"
//                       className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
//                     />
//                   </div>

//                   {currentLine.date && (
//                     <div className="col-span-2">
//                       <label className="block text-slate-400 mb-1 text-sm">Day</label>
//                       <input
//                         type="text"
//                         value={formatDateToDay(currentLine.date)}
//                         readOnly
//                         className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400 text-sm"
//                       />
//                     </div>
//                   )}
//                 </div>

//                 <div className="flex justify-end gap-3 pt-4 border-t border-slate-800">
//                   <button
//                     onClick={() => { setIsAddingLine(false); setSelectedLineIndex(null); }}
//                     className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg cursor-pointer"
//                   >
//                     Cancel
//                   </button>
//                   <button
//                     onClick={handleSaveLine}
//                     className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg cursor-pointer"
//                   >
//                     Save Line
//                   </button>
//                 </div>
//               </div>
//             </div>
//           </div>
//         )}
//       </div>
//     );
//   }

//   // Render halaman lain (jika perlu navigasi tambahan)
//   return null;
// }

import React, { useState, useEffect, useCallback, useMemo, useRef, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import {
  Building2, Briefcase, HardHat, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
  Pencil, Trash2, Lock, X, Check, Printer, FileCheck, Clock, Sparkles, 
  FileSpreadsheet, Users, Search, Plus, Calendar
} from 'lucide-react';

interface OutsourcingPageProps {
  user: UserSession;
  onLogout: () => void;
}

interface TimesheetLine {
  id: string;
  date: string;
  workOrder: string;
  description: string;
  effectiveHours: number;
  overtimeHours: number;
}

interface TimesheetHeader {
  id: string;
  code: string;
  startDate: string;
  endDate: string;
  responsible: string;
  unitKerja: string;
  status: string;
  lines: TimesheetLine[];
}

export default function OutsourcingPage({ user, onLogout }: OutsourcingPageProps) {
  const [viewMode, setViewMode] = useState<'list' | 'create'>('list');
  
  // State Form Header Timesheet
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [unitKerja, setUnitKerja] = useState('Biro Dukungan & Administrasi');
  
  // State Lines Timesheet
  const [lines, setLines] = useState<TimesheetLine[]>([]);
  
  // State Modal Add Line (Gambar Kedua)
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [lineDate, setLineDate] = useState('');
  const [lineWorkOrder, setLineWorkOrder] = useState('');
  const [lineDescription, setLineDescription] = useState('');
  const [lineEffective, setLineEffective] = useState('00:00');
  const [lineOvertime, setLineOvertime] = useState('00:00');

  // State daftar timesheet tersimpan
  const [savedTimesheets, setSavedTimesheets] = TimesheetListState();

  function TimesheetListState() {
    return useState<TimesheetHeader[]>([
      {
        id: '1',
        code: 'TIM-2609301065844',
        startDate: '2026-09-28',
        endDate: '2026-09-30',
        responsible: user.nama || 'Nur Zakiyyah',
        unitKerja: 'Biro Dukungan & Administrasi',
        status: 'Approved',
        lines: []
      }
    ]);
  }

  // Load data job cards / work orders yang diterbitkan untuk user ini
  const [availableWorkOrders, setAvailableWorkOrders] = useState<any[]>([]);

  useEffect(() => {
    async function fetchWorkOrders() {
      try {
        const { data, error } = await supabase.from('job_cards').select('*');
        if (data) {
          // Filter work order yang PIC-nya sesuai dengan nama user yang sedang login
          const myWo = data.filter(item => 
            (item.pic || '').toLowerCase() === (user.nama || '').toLowerCase() || true // fallback fleksibel
          );
          setAvailableWorkOrders(myWo);
        }
      } catch {}
    }
    fetchWorkOrders();
  }, [user.nama]);

  const handleAddLineSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!lineDate || !lineWorkOrder) {
      alert('Tanggal dan Work Order wajib diisi!');
      return;
    }

    const newLine: TimesheetLine = {
      id: Math.random().toString(36.substring(2, 9)),
      date: lineDate,
      workOrder: lineWorkOrder,
      description: lineDescription || 'Pekerjaan Desain & Drafting',
      effectiveHours: parseFloat(lineEffective) || 8,
      overtimeHours: parseFloat(lineOvertime) || 0,
    };

    setLines(prev => [...prev, newLine]);
    setIsModalOpen(false);
    // Reset modal form
    setLineWorkOrder('');
    setLineDescription('');
    setLineEffective('08:00');
    setLineOvertime('00:00');
  };

  const handleSaveTimesheet = () => {
    if (!startDate || !endDate) {
      alert('Harap tentukan Tanggal Mulai dan Tanggal Selesai terlebih dahulu.');
      return;
    }
    if (lines.length === 0) {
      alert('Tambahkan setidaknya satu baris timesheet (Add a line).');
      return;
    }

    const randomCode = `TIM-${Math.floor(1000000000000 + Math.random() * 9000000000000)}`;
    const newTimesheet: TimesheetHeader = {
      id: Date.now().toString(),
      code: randomCode,
      startDate,
      endDate,
      responsible: user.nama,
      unitKerja,
      status: 'Approved',
      lines
    };

    setSavedTimesheets(prev => [newTimesheet, ...prev]);
    setViewMode('list');
    setLines([]);
    setStartDate('');
    setEndDate('');
    alert(`Timesheet berhasil disimpan dengan nomor: ${randomCode}`);
  };

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-amber-600 rounded-lg text-white">
              <HardHat className="w-4 h-4" />
            </div>
            <span className="font-bold text-sm text-white">PORTAL MITRA / OUTSOURCING — <span className="text-amber-400">{user.nama}</span></span>
          </div>
          <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
            Keluar
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-6">
        
        {viewMode === 'list' ? (
          /* ================= VIEW 1: KARTU TIMESHEET (GAMBAR KETIGA) ================= */
          <div className="space-y-6">
            <div className="flex items-center justify-between border-b border-slate-800 pb-4">
              <div>
                <h2 className="text-xl font-bold text-white">Project Timesheet</h2>
                <span className="text-xs text-slate-400">Dokumen rekapitulasi jam kerja dan work order harian</span>
              </div>
              <button
                onClick={() => setViewMode('create')}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs rounded-lg shadow flex items-center gap-1.5 cursor-pointer"
              >
                <Plus className="w-4 h-4" /> Create
              </button>
            </div>

            {/* Grid Kartu Seperti Gambar Ketiga */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
              {savedTimesheets.map((ts) => (
                <div key={ts.id} className="bg-slate-900 border border-slate-800 hover:border-amber-500/50 rounded-xl p-4 space-y-3 transition shadow-lg">
                  <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                    <span className="font-mono font-bold text-amber-400 text-xs tracking-wider">{ts.code}</span>
                    <span className="text-[10px] px-2 py-0.5 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded font-semibold">
                      {ts.status}
                    </span>
                  </div>
                  <div className="text-xs space-y-1.5 font-mono text-slate-300">
                    <div><span className="text-slate-500">Responsible:</span> <span className="text-white">{ts.responsible}</span></div>
                    <div className="flex justify-between"><span className="text-slate-500">Start Date:</span> <span className="text-cyan-300">{ts.startDate}</span></div>
                    <div className="flex justify-between"><span className="text-slate-500">End Date:</span> <span className="text-cyan-300">{ts.endDate}</span></div>
                  </div>
                  <div className="pt-2 border-t border-slate-800/80 text-[11px] text-slate-400 flex justify-between">
                    <span>Total Baris: {ts.lines.length} Line(s)</span>
                    <span className="text-purple-400 font-semibold">{ts.unitKerja}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          /* ================= VIEW 2: FORM CREATE TIMESHEET (GAMBAR PERTAMA & KEDUA) ================= */
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <button onClick={() => setViewMode('list')} className="p-1.5 bg-slate-800 hover:bg-slate-700 rounded-lg text-slate-300 cursor-pointer">
                  <ArrowLeft className="w-4 h-4" />
                </button>
                <h2 className="text-lg font-bold text-white">New Timesheet</h2>
              </div>
              <button
                onClick={handleSaveTimesheet}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-lg cursor-pointer"
              >
                Simpan & Terbitkan
              </button>
            </div>

            {/* Header Informasi (Gambar Pertama) */}[cite: 9]
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs bg-slate-950 p-4 rounded-xl border border-slate-800">
              <div className="space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">User</span>
                  <span className="font-semibold text-white">{user.nama}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Date Start</span>
                  <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} style={{ colorScheme: 'dark' }}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono cursor-pointer" />
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Date End</span>
                  <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} style={{ colorScheme: 'dark' }}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono cursor-pointer" />
                </div>
              </div>

              <div className="space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Unit Kerja</span>
                  <span className="font-semibold text-cyan-400">{unitKerja}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Divisi</span>
                  <span className="font-semibold text-white">71000 - Divisi Desain</span>
                </div>
              </div>
            </div>

            {/* Tab Project Timesheet Line (Gambar Pertama) */}[cite: 9]
            <div className="space-y-3">
              <div className="border-b border-slate-800 pb-2">
                <span className="text-xs font-bold text-slate-300 uppercase tracking-wider bg-slate-800 px-3 py-1.5 rounded-t-lg">
                  Project Timesheet Line
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="text-slate-400 border-b border-slate-800 font-mono">
                      <th className="py-2.5 px-3">Date</th>
                      <th className="py-2.5 px-3">Work Order</th>
                      <th className="py-2.5 px-3">Description</th>
                      <th className="py-2.5 px-3 text-center">Effective Hours</th>
                      <th className="py-2.5 px-3 text-center">Overtime Hours</th>
                      <th className="py-2.5 px-3 text-center">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-slate-300">
                    {lines.map((l, idx) => (
                      <tr key={l.id} className="hover:bg-slate-950/50">
                        <td className="py-2.5 px-3 font-mono text-cyan-300">{l.date}</td>
                        <td className="py-2.5 px-3 font-mono font-bold text-amber-300">{l.workOrder}</td>
                        <td className="py-2.5 px-3 text-slate-200">{l.description}</td>
                        <td className="py-2.5 px-3 text-center font-mono text-emerald-400">{l.effectiveHours} Jam</td>
                        <td className="py-2.5 px-3 text-center font-mono text-purple-400">{l.overtimeHours} Jam</td>
                        <td className="py-2.5 px-3 text-center">
                          <button onClick={() => setLines(lines.filter(item => item.id !== l.id))} className="text-rose-400 hover:text-rose-300 cursor-pointer">
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Tombol Add a Line (Gambar Pertama) */}[cite: 9]
              <button
                onClick={() => {
                  if (!startDate || !endDate) {
                    alert('Mohon isi Date Start dan Date End terlebih dahulu!');
                    return;
                  }
                  setLineDate(startDate);
                  setIsModalOpen(true);
                }}
                className="text-blue-400 hover:text-blue-300 text-xs font-semibold cursor-pointer pt-2 flex items-center gap-1"
              >
                + Add a line
              </button>
            </div>
          </div>
        )}
      </main>

      {/* ================= MODAL CREATE LINES (GAMBAR KEDUA) ================= */}[cite: 10]
      {isModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
            <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <span className="font-bold text-xs text-white">Create Lines (Timesheet Harian)</span>
              <button onClick={() => setIsModalOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleAddLineSubmit} className="p-5 space-y-4 text-xs">
              <div className="grid grid-cols-1 gap-3">
                <div>
                  <label className="block text-slate-400 mb-1">Date (Tanggal Harian)</label>
                  <input
                    type="date"
                    min={startDate}
                    max={endDate}
                    value={lineDate}
                    onChange={(e) => setLineDate(e.target.value)}
                    required
                    style={{ colorScheme: 'dark' }}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono cursor-pointer"
                  />
                  <span className="text-[10px] text-slate-500 mt-0.5 block">Harus berada di antara rentang Start Date & End Date.</span>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1">Work Order (Terbitan Kabiro)</label>
                  <select
                    value={lineWorkOrder}
                    onChange={(e) => {
                      setLineWorkOrder(e.target.value);
                      const found = availableWorkOrders.find(w => (w.kode_jc || w.packageTitle) === e.target.value);
                      if (found) setLineDescription(found.task_name || '');
                    }}
                    required
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                  >
                    <option value="">-- Pilih Work Order --</option>
                    {availableWorkOrders.map((wo, idx) => (
                      <option key={idx} value={wo.kode_jc || `WO-${idx}`}>
                        {wo.kode_jc || `WO-${idx}`} - {wo.task_name || wo.project}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1">Description</label>
                  <input
                    type="text"
                    value={lineDescription}
                    onChange={(e) => setLineDescription(e.target.value)}
                    placeholder="Uraian pekerjaan harian..."
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                  />
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-slate-400 mb-1">Effective Hours (Jam)</label>
                    <input
                      type="number"
                      step="0.5"
                      min="0"
                      max="24"
                      value={lineEffective}
                      onChange={(e) => setLineEffective(e.target.value)}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-400 mb-1">Overtime Hours (Jam Lembur)</label>
                    <input
                      type="number"
                      step="0.5"
                      min="0"
                      max="12"
                      value={lineOvertime}
                      onChange={(e) => setLineOvertime(e.target.value)}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                    />
                  </div>
                </div>
              </div>

              <div className="pt-3 border-t border-slate-800 flex justify-end gap-2">
                <button type="button" onClick={() => setIsModalOpen(false)} className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-semibold rounded-lg cursor-pointer">
                  Discard
                </button>
                <button type="submit" className="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg cursor-pointer">
                  Save & Close
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}