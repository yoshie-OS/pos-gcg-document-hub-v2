import React, { useState, useEffect } from 'react';
import Sidebar from '@/components/layout/Sidebar';
import Topbar from '@/components/layout/Topbar';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Textarea } from '@/components/ui/textarea';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Badge } from '@/components/ui/badge';
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { useSidebar } from '@/contexts/SidebarContext';
import { useYear } from '@/contexts/YearContext';
import { useAOI } from '@/contexts/AOIContext';
import { useStrukturPerusahaan } from '@/contexts/StrukturPerusahaanContext';
import { useToast } from '@/hooks/use-toast';
import { PageHeaderPanel, YearSelectorPanel } from '@/components/panels';
import {
  Plus,
  Edit,
  Trash2,
  Star,
  Users,
  FileText,
  CheckCircle,
  ChevronDown,
  ChevronRight,
  Save,
  X,
  Upload,
  RefreshCw
} from 'lucide-react';

const AOIManagement = () => {
  const { isSidebarOpen } = useSidebar();
  const { selectedYear, setSelectedYear, availableYears } = useYear();
  const {
    aoiTables,
    aoiRecommendations,
    createAOITable,
    updateAOITable,
    deleteAOITable,
    addRecommendation,
    updateRecommendation,
    deleteRecommendation
  } = useAOI();
  const { direktorat, subdirektorat, divisi } = useStrukturPerusahaan();
  const { toast } = useToast();

  // Get current user info
  const [currentUser, setCurrentUser] = useState<any>(null);

  useEffect(() => {
    const userStr = localStorage.getItem('user');
    if (userStr) {
      try {
        const user = JSON.parse(userStr);
        setCurrentUser(user);
      } catch (error) {
        console.error('Error parsing user:', error);
      }
    }
  }, []);

  // State
  const [expandedTables, setExpandedTables] = useState<Set<number>>(new Set());
  const [newTableName, setNewTableName] = useState('');
  const [editingRowId, setEditingRowId] = useState<number | null>(null);

  // Inline add form for Rekomendasi per table
  const [inlineRekomendasiForm, setInlineRekomendasiForm] = useState<{[key: number]: any}>({});
  const [isAddingRekomendasi, setIsAddingRekomendasi] = useState<{[key: number]: boolean}>({});

  // Inline add form for Saran per table
  const [inlineSaranForm, setInlineSaranForm] = useState<{[key: number]: any}>({});
  const [isAddingSaran, setIsAddingSaran] = useState<{[key: number]: boolean}>({});

  // Edit row form state
  const [editRowForm, setEditRowForm] = useState<any>({});

  // Org form state per table
  const [orgForms, setOrgForms] = useState<{[key: number]: any}>({});

  // Org level selection per table (direktorat/subdirektorat/divisi)
  const [orgLevels, setOrgLevels] = useState<{[key: number]: string}>({});

  // Initialize org forms and inline add forms when tables are loaded
  useEffect(() => {
    if (aoiTables && aoiTables.length > 0) {
      aoiTables.forEach(table => {
        // Initialize org forms
        if (!orgForms[table.id]) {
          setOrgForms(prev => ({
            ...prev,
            [table.id]: {
              targetDirektorat: table.targetDirektorat || '',
              targetSubdirektorat: table.targetSubdirektorat || '',
              targetDivisi: table.targetDivisi || ''
            }
          }));
        }
        if (!orgLevels[table.id]) {
          const currentLevel = (table.targetDivisi && table.targetDivisi.trim())
            ? 'divisi'
            : (table.targetSubdirektorat && table.targetSubdirektorat.trim())
              ? 'subdirektorat'
              : 'direktorat';
          setOrgLevels(prev => ({
            ...prev,
            [table.id]: currentLevel
          }));
        }

        // Initialize inline forms
        initializeInlineForms(table.id);
      });
    }
  }, [aoiTables]);

  // Filter tables berdasarkan tahun dan struktur organisasi user
  const yearTables = selectedYear ? aoiTables.filter(table => {
    console.log(`[AOI Filter] Checking table ID ${table.id}, year ${table.tahun} vs selected ${selectedYear}`);
    console.log(`[AOI Filter] Table targets:`, {
      direktorat: table.targetDirektorat,
      subdirektorat: table.targetSubdirektorat,
      divisi: table.targetDivisi
    });

    if (table.tahun !== selectedYear) {
      console.log(`[AOI Filter] ❌ Year mismatch, skipping`);
      return false;
    }

    // Super admin can see all
    if (currentUser?.role === 'superadmin' || currentUser?.role === 'super-admin') {
      console.log(`[AOI Filter] ✅ Super admin - showing all`);
      return true;
    }

    // Regular user - filter by structure
    if (!currentUser?.subdirektorat) return false;

    const userSubdirektorat = currentUser.subdirektorat;
    const userDivisi = currentUser.divisi || '';

    // Match by divisi (most specific)
    if (table.targetDivisi && table.targetDivisi.trim()) {
      return table.targetDivisi === userDivisi;
    }

    // Match by subdirektorat
    if (table.targetSubdirektorat && table.targetSubdirektorat.trim()) {
      return table.targetSubdirektorat === userSubdirektorat;
    }

    // Match by direktorat
    if (table.targetDirektorat && table.targetDirektorat.trim()) {
      // Find user's subdirektorat data
      const userSubdirektoratData = subdirektorat.find(
        sub => sub.nama === userSubdirektorat && sub.tahun === selectedYear
      );

      if (!userSubdirektoratData) {
        // If subdirektorat data not found, user can't see direktorat-level tables
        console.warn(`Subdirektorat "${userSubdirektorat}" not found for year ${selectedYear}`);
        return false;
      }

      // Find the direktorat this subdirektorat belongs to
      const userDirektoratData = direktorat.find(
        dir => dir.id === userSubdirektoratData.direktoratId && dir.tahun === selectedYear
      );

      if (!userDirektoratData) {
        // If direktorat data not found, user can't see direktorat-level tables
        console.warn(`Direktorat with ID ${userSubdirektoratData.direktoratId} not found for year ${selectedYear}`);
        return false;
      }

      // Match the direktorat name
      const matches = userDirektoratData.nama === table.targetDirektorat;
      console.log(`Direktorat match: User="${userDirektoratData.nama}" vs Table="${table.targetDirektorat}" = ${matches}`);
      return matches;
    }

    // If no organization assigned, show to all users
    return true;
  }) : [];

  // Filter data struktur perusahaan berdasarkan tahun yang dipilih
  const yearDirektorat = selectedYear ? direktorat.filter(dir => dir.tahun === selectedYear) : [];
  const yearSubdirektorat = selectedYear ? subdirektorat.filter(sub => sub.tahun === selectedYear) : [];
  const yearDivisi = selectedYear ? divisi.filter(div => div.tahun === selectedYear) : [];

  // Debug: Check if organizational data is loaded
  useEffect(() => {
    console.log('=== AOIManagement Organizational Data Debug ===');
    console.log('Selected Year:', selectedYear);
    console.log('All Direktorat:', direktorat);
    console.log('All Subdirektorat:', subdirektorat);
    console.log('All Divisi:', divisi);
    console.log('Filtered yearDirektorat:', yearDirektorat);
    console.log('Filtered yearSubdirektorat:', yearSubdirektorat);
    console.log('Filtered yearDivisi:', yearDivisi);
    console.log('============================================');
  }, [selectedYear, direktorat, subdirektorat, divisi, yearDirektorat, yearSubdirektorat, yearDivisi]);

  // Get subdirektorat berdasarkan direktorat yang dipilih
  const getSubdirektoratByDirektorat = (direktoratId: number) => {
    return yearSubdirektorat.filter(sub => sub.direktoratId === direktoratId);
  };

  // Get divisi berdasarkan subdirektorat yang dipilih
  const getDivisiBySubdirektorat = (subdirektoratId: number) => {
    return yearDivisi.filter(div => div.subdirektoratId === subdirektoratId);
  };

  // Handle create table
  const handleCreateTable = (e: React.FormEvent) => {
    e.preventDefault();

    if (!selectedYear) {
      toast({
        title: "Data tidak lengkap",
        description: "Pilih tahun untuk AOI",
        variant: "destructive"
      });
      return;
    }

    if (!newTableName.trim()) {
      toast({
        title: "Data tidak lengkap",
        description: "Nama tabel AOI wajib diisi",
        variant: "destructive"
      });
      return;
    }

    const payload = {
      nama: newTableName.trim(),
      deskripsi: '',
      tahun: selectedYear,
      status: 'active' as 'active' | 'inactive',
      targetType: 'direktorat' as 'direktorat' | 'subdirektorat' | 'divisi',
      targetDirektorat: '',
      targetSubdirektorat: '',
      targetDivisi: ''
    };

    createAOITable(payload as any);
    toast({
      title: "Tabel AOI berhasil dibuat",
      description: "Tabel baru telah ditambahkan"
    });

    setNewTableName('');
  };

  // Toggle table expansion
  const toggleTableExpansion = (tableId: number) => {
    const newExpanded = new Set(expandedTables);
    if (newExpanded.has(tableId)) {
      newExpanded.delete(tableId);
    } else {
      newExpanded.add(tableId);
    }
    setExpandedTables(newExpanded);
  };

  // Render star rating
  const renderStars = (rating: string) => {
    if (rating === 'TIDAK_ADA' || !rating) {
      return <span className="text-xs text-gray-400">-</span>;
    }

    const ratingMap: Record<string, number> = {
      'RENDAH': 1,
      'SEDANG': 2,
      'TINGGI': 3,
      'SANGAT_TINGGI': 4,
      'KRITIS': 5
    };
    const starCount = ratingMap[rating] || 0;

    return (
      <div className="flex justify-center">
        {Array.from({ length: 5 }, (_, i) => (
          <Star
            key={i}
            className={`w-3 h-3 ${
              i < starCount ? 'text-yellow-500 fill-current' : 'text-gray-300'
            }`}
          />
        ))}
      </div>
    );
  };

  // Get recommendations for a specific table
  const getTableRecommendations = (tableId: number) => {
    return (aoiRecommendations || []).filter(rec => rec.aoiTableId === tableId);
  };

  // Get REKOMENDASI for a specific table
  const getTableRekomendasi = (tableId: number) => {
    return (aoiRecommendations || []).filter(rec => rec.aoiTableId === tableId && rec.jenis === 'REKOMENDASI');
  };

  // Get SARAN for a specific table
  const getTableSaran = (tableId: number) => {
    return (aoiRecommendations || []).filter(rec => rec.aoiTableId === tableId && rec.jenis === 'SARAN');
  };

  // Get next recommendation number for a table and jenis
  const getNextRecommendationNumber = (tableId: number, jenis: 'REKOMENDASI' | 'SARAN') => {
    const tableRecs = jenis === 'REKOMENDASI'
      ? getTableRekomendasi(tableId)
      : getTableSaran(tableId);
    return tableRecs.length + 1;
  };

  // Initialize inline forms for a table
  const initializeInlineForms = (tableId: number) => {
    if (!inlineRekomendasiForm[tableId]) {
      setInlineRekomendasiForm({
        ...inlineRekomendasiForm,
        [tableId]: {
          isi: '',
          aspekAOI: '',
          tingkatUrgensi: 'TIDAK_ADA',
          pihakTerkait: '',
          organPerusahaan: ''
        }
      });
    }
    if (!inlineSaranForm[tableId]) {
      setInlineSaranForm({
        ...inlineSaranForm,
        [tableId]: {
          isi: '',
          aspekAOI: '',
          tingkatUrgensi: 'TIDAK_ADA',
          pihakTerkait: '',
          organPerusahaan: ''
        }
      });
    }
  };

  // Handle add REKOMENDASI
  const handleAddRekomendasi = async (tableId: number) => {
    const form = inlineRekomendasiForm[tableId];
    if (!form || !form.isi.trim()) {
      toast({
        title: "Data tidak lengkap",
        description: "Isi rekomendasi wajib diisi",
        variant: "destructive"
      });
      return;
    }

    setIsAddingRekomendasi({...isAddingRekomendasi, [tableId]: true});
    try {
      const payload = {
        no: getNextRecommendationNumber(tableId, 'REKOMENDASI'),
        aoiTableId: tableId,
        jenis: 'REKOMENDASI' as 'REKOMENDASI' | 'SARAN',
        isi: form.isi,
        tingkatUrgensi: form.tingkatUrgensi || 'TIDAK_ADA',
        aspekAOI: form.aspekAOI || '',
        pihakTerkait: form.pihakTerkait || '',
        organPerusahaan: form.organPerusahaan || '',
        status: 'active' as 'active' | 'inactive'
      };

      await addRecommendation(payload);
      toast({
        title: "Rekomendasi berhasil ditambahkan",
        description: "Data telah disimpan"
      });

      // Reset form
      setInlineRekomendasiForm({
        ...inlineRekomendasiForm,
        [tableId]: {
          isi: '',
          aspekAOI: '',
          tingkatUrgensi: 'TIDAK_ADA',
          pihakTerkait: '',
          organPerusahaan: ''
        }
      });
    } finally {
      setIsAddingRekomendasi({...isAddingRekomendasi, [tableId]: false});
    }
  };

  // Handle add SARAN
  const handleAddSaran = async (tableId: number) => {
    const form = inlineSaranForm[tableId];
    if (!form || !form.isi.trim()) {
      toast({
        title: "Data tidak lengkap",
        description: "Isi saran wajib diisi",
        variant: "destructive"
      });
      return;
    }

    setIsAddingSaran({...isAddingSaran, [tableId]: true});
    try {
      const payload = {
        no: getNextRecommendationNumber(tableId, 'SARAN'),
        aoiTableId: tableId,
        jenis: 'SARAN' as 'REKOMENDASI' | 'SARAN',
        isi: form.isi,
        tingkatUrgensi: form.tingkatUrgensi || 'TIDAK_ADA',
        aspekAOI: form.aspekAOI || '',
        pihakTerkait: form.pihakTerkait || '',
        organPerusahaan: form.organPerusahaan || '',
        status: 'active' as 'active' | 'inactive'
      };

      await addRecommendation(payload);
      toast({
        title: "Saran berhasil ditambahkan",
        description: "Data telah disimpan"
      });

      // Reset form
      setInlineSaranForm({
        ...inlineSaranForm,
        [tableId]: {
          isi: '',
          aspekAOI: '',
          tingkatUrgensi: 'TIDAK_ADA',
          pihakTerkait: '',
          organPerusahaan: ''
        }
      });
    } finally {
      setIsAddingSaran({...isAddingSaran, [tableId]: false});
    }
  };

  // Start editing row
  const startEditingRow = (rec: any) => {
    setEditingRowId(rec.id);
    setEditRowForm({
      jenis: rec.jenis,
      isi: rec.isi,
      aspekAOI: rec.aspekAOI || '',
      tingkatUrgensi: rec.tingkatUrgensi || 'TIDAK_ADA',
      pihakTerkait: rec.pihakTerkait || '',
      organPerusahaan: rec.organPerusahaan || ''
    });
  };

  // Cancel editing row
  const cancelEditingRow = () => {
    setEditingRowId(null);
    setEditRowForm({});
  };

  // Save edited row
  const saveEditedRow = async (recId: number) => {
    if (!editRowForm.isi.trim()) {
      toast({
        title: "Data tidak lengkap",
        description: "Isi rekomendasi wajib diisi",
        variant: "destructive"
      });
      return;
    }

    await updateRecommendation(recId, editRowForm);
    toast({
      title: "Rekomendasi berhasil diupdate",
      description: "Data telah diperbarui"
    });

    cancelEditingRow();
  };

  // Delete row
  const deleteRow = async (recId: number) => {
    if (confirm('Apakah Anda yakin ingin menghapus rekomendasi ini?')) {
      await deleteRecommendation(recId);
      toast({
        title: "Rekomendasi berhasil dihapus",
        description: "Data telah dihapus"
      });
    }
  };

  // Delete table
  const handleDeleteTable = (tableId: number) => {
    if (confirm('Apakah Anda yakin ingin menghapus tabel ini? Semua rekomendasi akan ikut terhapus.')) {
      deleteAOITable(tableId);
      toast({
        title: "Tabel AOI berhasil dihapus",
        description: "Data telah dihapus"
      });
    }
  };

  // Save org settings for a table
  const saveOrgSettingsInline = async (tableId: number) => {
    const form = orgForms[tableId];
    const level = orgLevels[tableId] || 'direktorat';

    if (!form) {
      toast({
        title: "Error",
        description: "Form data tidak ditemukan",
        variant: "destructive"
      });
      return;
    }

    await updateAOITable(tableId, {
      targetType: level as 'direktorat' | 'subdirektorat' | 'divisi',
      targetDirektorat: form.targetDirektorat,
      targetSubdirektorat: form.targetSubdirektorat,
      targetDivisi: form.targetDivisi
    });

    toast({
      title: "Target organisasi berhasil diupdate",
      description: "Pengaturan telah disimpan"
    });
  };

  if (!selectedYear) {
    return (
      <>
        <Sidebar />
        <Topbar />
        <div className={`transition-all duration-300 ease-in-out pt-16 ${isSidebarOpen ? 'lg:ml-64' : 'ml-0'}`}>
          <div className="p-6">
            <div className="text-center py-20 bg-white/80 backdrop-blur-sm rounded-2xl shadow-xl border border-white/20">
              <div className="relative">
                <div className="absolute inset-0 bg-gradient-to-r from-blue-400 to-purple-400 rounded-full blur-3xl opacity-20 animate-pulse"></div>
                <div className="relative z-10">
                  <div className="w-20 h-20 bg-gradient-to-r from-blue-500 to-purple-500 rounded-full flex items-center justify-center mx-auto mb-6">
                    <FileText className="w-10 h-10 text-white" />
                  </div>
                  <h3 className="text-2xl font-bold text-gray-900 mb-4 bg-gradient-to-r from-blue-600 to-purple-600 bg-clip-text text-transparent">
                    Pilih Tahun Buku
                  </h3>
                  <p className="text-gray-600 text-lg max-w-md mx-auto">
                    Silakan pilih tahun buku untuk mengelola Area of Improvement (AOI)
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </>
    );
  }

  return (
    <>
      <Sidebar />
      <Topbar />

      <div className={`transition-all duration-300 ease-in-out pt-16 ${isSidebarOpen ? 'lg:ml-64' : 'ml-0'}`}>
        <div className="p-6">
          {/* Header */}
          <YearSelectorPanel
            className="mb-4"
            selectedYear={selectedYear}
            onYearChange={setSelectedYear}
            availableYears={availableYears}
            title="Tahun Buku"
            description="Pilih tahun buku untuk mengelola AOI"
            data-tour="year-selector"
          />
          <PageHeaderPanel
            title="Area of Improvement (AOI) Management"
            subtitle={`Kelola rekomendasi perbaikan GCG untuk tahun ${selectedYear}`}
          />

          {/* Create AOI Table Form - Simplified */}
          <Card className="mb-6 border-2 border-blue-200 shadow-lg bg-gradient-to-r from-white to-blue-50" data-tour="create-table">
            <CardHeader className="pb-4">
              <CardTitle className="flex items-center space-x-2 text-blue-900">
                <Plus className="w-5 h-5 text-blue-600" />
                <span>Buat Tabel AOI Baru</span>
              </CardTitle>
              <p className="text-sm text-gray-600">
                Buat tabel AOI untuk mengelola rekomendasi perbaikan GCG
              </p>
            </CardHeader>
            <CardContent>
              <form onSubmit={handleCreateTable} className="space-y-4">
                <div className="flex gap-2">
                  <div className="flex-1">
                    <Input
                      type="text"
                      value={newTableName}
                      onChange={(e) => setNewTableName(e.target.value)}
                      className="border-gray-300 focus:border-blue-500"
                      placeholder="Masukkan nama tabel AOI..."
                      required
                    />
                  </div>
                  <Button type="submit" className="bg-blue-600 hover:bg-blue-700">
                    <Plus className="w-4 h-4 mr-2" />
                    Buat Tabel
                  </Button>
                </div>
              </form>
            </CardContent>
          </Card>

          {/* AOI Tables */}
          <div className="space-y-6">
            {yearTables.map((table) => (
              <Card key={table.id} className="border-0 shadow-lg bg-gradient-to-r from-white to-blue-50">
                <CardHeader className="pb-4">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-3 flex-1">
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => toggleTableExpansion(table.id)}
                        className="p-1 h-8 w-8"
                        data-tour={table.id === yearTables[0]?.id ? "expand-table" : undefined}
                      >
                        {expandedTables.has(table.id) ? (
                          <ChevronDown className="w-4 h-4" />
                        ) : (
                          <ChevronRight className="w-4 h-4" />
                        )}
                      </Button>
                      <div className="flex-1">
                        <CardTitle className="flex flex-wrap items-center gap-3 text-blue-900">
                          <div className="flex items-center space-x-2">
                            <FileText className="w-5 h-5 text-blue-600" />
                            <span>{table.nama}</span>
                          </div>
                          <div className="flex items-center gap-2">
                            {table.targetDirektorat && table.targetDirektorat.trim() ? (
                              <Badge variant="outline" className="bg-blue-100 text-blue-800 border-blue-200">
                                {table.targetDirektorat}
                              </Badge>
                            ) : null}
                            {table.targetSubdirektorat && table.targetSubdirektorat.trim() ? (
                              <Badge variant="outline" className="bg-green-100 text-green-800 border-green-200">
                                {table.targetSubdirektorat}
                              </Badge>
                            ) : null}
                            {table.targetDivisi && table.targetDivisi.trim() ? (
                              <Badge variant="outline" className="bg-purple-100 text-purple-800 border-purple-200">
                                {table.targetDivisi}
                              </Badge>
                            ) : null}
                          </div>
                        </CardTitle>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() => handleDeleteTable(table.id)}
                        className="border-red-200 text-red-600 hover:bg-red-50"
                        data-tour={table.id === yearTables[0]?.id ? "delete-table" : undefined}
                      >
                        <Trash2 className="w-4 h-4 mr-2" />
                        Hapus
                      </Button>
                    </div>
                  </div>
                </CardHeader>

                {expandedTables.has(table.id) && (
                  <CardContent>
                    {/* Target Organisasi Form */}
                    <div className="mb-6 p-4 bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-200 rounded-lg" data-tour="target-org">
                      <h4 className="text-sm font-semibold text-blue-900 mb-3">Target Organisasi</h4>
                      <div className="grid grid-cols-1 gap-4">
                        {/* Organizational Level Toggle */}
                        <div className="flex gap-2">
                          <Button
                            size="sm"
                            variant={orgLevels[table.id] === 'direktorat' ? 'default' : 'outline'}
                            onClick={() => setOrgLevels({...orgLevels, [table.id]: 'direktorat'})}
                            className={orgLevels[table.id] === 'direktorat' ? 'bg-blue-600' : 'border-blue-200 text-blue-600'}
                          >
                            Direktorat
                          </Button>
                          <Button
                            size="sm"
                            variant={orgLevels[table.id] === 'subdirektorat' ? 'default' : 'outline'}
                            onClick={() => setOrgLevels({...orgLevels, [table.id]: 'subdirektorat'})}
                            className={orgLevels[table.id] === 'subdirektorat' ? 'bg-blue-600' : 'border-blue-200 text-blue-600'}
                          >
                            Subdirektorat
                          </Button>
                          <Button
                            size="sm"
                            variant={orgLevels[table.id] === 'divisi' ? 'default' : 'outline'}
                            onClick={() => setOrgLevels({...orgLevels, [table.id]: 'divisi'})}
                            className={orgLevels[table.id] === 'divisi' ? 'bg-blue-600' : 'border-blue-200 text-blue-600'}
                          >
                            Divisi
                          </Button>
                        </div>

                        {/* Organizational Structure Dropdowns - Conditional based on selected level */}
                        <div className="grid grid-cols-1 gap-3">
                          {/* Always show Direktorat for all levels */}
                          <div>
                            <Label className="text-xs font-medium text-gray-700 mb-1">Direktorat</Label>
                            <Select
                              value={yearDirektorat.find(d => d.nama === orgForms[table.id]?.targetDirektorat)?.id.toString() || 'empty'}
                              onValueChange={(value) => {
                                if (value === 'empty') {
                                  setOrgForms({
                                    ...orgForms,
                                    [table.id]: {...orgForms[table.id], targetDirektorat: '', targetSubdirektorat: '', targetDivisi: ''}
                                  });
                                } else {
                                  const d = yearDirektorat.find(x => x.id.toString() === value);
                                  setOrgForms({
                                    ...orgForms,
                                    [table.id]: {...orgForms[table.id], targetDirektorat: d?.nama || '', targetSubdirektorat: '', targetDivisi: ''}
                                  });
                                }
                              }}
                            >
                              <SelectTrigger className="h-9 text-xs">
                                <SelectValue placeholder="-- Pilih Direktorat --" />
                              </SelectTrigger>
                              <SelectContent>
                                <SelectItem value="empty">
                                  <span className="text-gray-500">-- Belum ditentukan --</span>
                                </SelectItem>
                                {yearDirektorat.map(d => (
                                  <SelectItem key={d.id} value={d.id.toString()}>{d.nama}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </div>

                          {/* Show Subdirektorat only for subdirektorat or divisi level */}
                          {(orgLevels[table.id] === 'subdirektorat' || orgLevels[table.id] === 'divisi') && (
                            <div>
                              <Label className="text-xs font-medium text-gray-700 mb-1">Subdirektorat</Label>
                              <Select
                                disabled={!orgForms[table.id]?.targetDirektorat}
                                value={yearSubdirektorat.find(s => s.nama === orgForms[table.id]?.targetSubdirektorat)?.id.toString() || 'empty'}
                                onValueChange={(value) => {
                                  if (value === 'empty') {
                                    setOrgForms({
                                      ...orgForms,
                                      [table.id]: {...orgForms[table.id], targetSubdirektorat: '', targetDivisi: ''}
                                    });
                                  } else {
                                    const s = yearSubdirektorat.find(x => x.id.toString() === value);
                                    setOrgForms({
                                      ...orgForms,
                                      [table.id]: {...orgForms[table.id], targetSubdirektorat: s?.nama || '', targetDivisi: ''}
                                    });
                                  }
                                }}
                              >
                                <SelectTrigger className="h-9 text-xs">
                                  <SelectValue placeholder="-- Pilih Subdirektorat --" />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="empty">
                                    <span className="text-gray-500">-- Belum ditentukan --</span>
                                  </SelectItem>
                                  {getSubdirektoratByDirektorat(
                                    yearDirektorat.find(d => d.nama === orgForms[table.id]?.targetDirektorat)?.id || 0
                                  ).map(s => (
                                    <SelectItem key={s.id} value={s.id.toString()}>{s.nama}</SelectItem>
                                  ))}
                                </SelectContent>
                              </Select>
                            </div>
                          )}

                          {/* Show Divisi only for divisi level */}
                          {orgLevels[table.id] === 'divisi' && (
                            <div>
                              <Label className="text-xs font-medium text-gray-700 mb-1">Divisi</Label>
                              <Select
                                disabled={!orgForms[table.id]?.targetSubdirektorat}
                                value={yearDivisi.find(v => v.nama === orgForms[table.id]?.targetDivisi)?.id.toString() || 'empty'}
                                onValueChange={(value) => {
                                  if (value === 'empty') {
                                    setOrgForms({
                                      ...orgForms,
                                      [table.id]: {...orgForms[table.id], targetDivisi: ''}
                                    });
                                  } else {
                                    const v = yearDivisi.find(x => x.id.toString() === value);
                                    setOrgForms({
                                      ...orgForms,
                                      [table.id]: {...orgForms[table.id], targetDivisi: v?.nama || ''}
                                    });
                                  }
                                }}
                              >
                                <SelectTrigger className="h-9 text-xs">
                                  <SelectValue placeholder="-- Pilih Divisi --" />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="empty">
                                    <span className="text-gray-500">-- Belum ditentukan --</span>
                                  </SelectItem>
                                  {getDivisiBySubdirektorat(
                                    yearSubdirektorat.find(s => s.nama === orgForms[table.id]?.targetSubdirektorat)?.id || 0
                                  ).map(v => (
                                    <SelectItem key={v.id} value={v.id.toString()}>{v.nama}</SelectItem>
                                  ))}
                                </SelectContent>
                              </Select>
                            </div>
                          )}
                        </div>

                        {/* Upload/Save Button */}
                        <div className="flex justify-end">
                          <Button
                            size="sm"
                            onClick={() => saveOrgSettingsInline(table.id)}
                            className="bg-blue-600 hover:bg-blue-700"
                          >
                            <Upload className="w-4 h-4 mr-2" />
                            Simpan Target Organisasi
                          </Button>
                        </div>
                      </div>
                    </div>

                    {/* Daftar Rekomendasi */}
                    <div className="mb-6" data-tour="rekomendasi-table">
                      <h4 className="text-lg font-semibold text-blue-900 mb-3">
                        Daftar Rekomendasi ({getTableRekomendasi(table.id).length})
                      </h4>
                      <div className="border border-blue-200 rounded-lg overflow-hidden bg-white shadow-sm">
                        <div className="overflow-x-auto">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-blue-50">
                                <TableHead className="text-blue-900 font-semibold w-12 text-center">NO</TableHead>
                                <TableHead className="text-blue-900 font-semibold min-w-[300px]">DESKRIPSI</TableHead>
                                <TableHead className="text-blue-900 font-semibold w-32">ASPEK AOI</TableHead>
                                <TableHead className="text-blue-900 font-semibold w-32 text-center">URGENSI</TableHead>
                                <TableHead className="text-blue-900 font-semibold w-40">PIHAK TERKAIT</TableHead>
                                <TableHead className="text-blue-900 font-semibold w-40">ORGAN PERUSAHAAN</TableHead>
                                <TableHead className="text-blue-900 font-semibold w-32 text-center">AKSI</TableHead>
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {/* Existing Rekomendasi rows */}
                              {getTableRekomendasi(table.id).map((rec, recIndex) => (
                                <TableRow key={rec.id} className="hover:bg-blue-50/50 border-b border-blue-100">
                                  <TableCell className="font-medium text-center text-blue-900">{rec.no}</TableCell>
                                  <TableCell>
                                  {editingRowId === rec.id ? (
                                    <Textarea
                                      value={editRowForm.isi}
                                      onChange={(e) => setEditRowForm({...editRowForm, isi: e.target.value})}
                                      className="text-xs min-h-[60px]"
                                      placeholder="Deskripsi rekomendasi..."
                                    />
                                  ) : (
                                    <div className="text-sm leading-relaxed text-gray-800">{rec.isi || '-'}</div>
                                  )}
                                </TableCell>
                                <TableCell>
                                  {editingRowId === rec.id ? (
                                    <Input
                                      value={editRowForm.aspekAOI}
                                      onChange={(e) => setEditRowForm({...editRowForm, aspekAOI: e.target.value})}
                                      className="h-8 text-xs"
                                      placeholder="Aspek AOI..."
                                    />
                                  ) : (
                                    <div className="text-xs">{rec.aspekAOI || '-'}</div>
                                  )}
                                </TableCell>
                                <TableCell className="text-center">
                                  {editingRowId === rec.id ? (
                                    <Select
                                      value={editRowForm.tingkatUrgensi}
                                      onValueChange={(v) => setEditRowForm({...editRowForm, tingkatUrgensi: v})}
                                    >
                                      <SelectTrigger className="h-8 text-xs">
                                        <SelectValue />
                                      </SelectTrigger>
                                      <SelectContent>
                                        <SelectItem value="TIDAK_ADA">-</SelectItem>
                                        <SelectItem value="RENDAH">⭐ Rendah</SelectItem>
                                        <SelectItem value="SEDANG">⭐⭐ Sedang</SelectItem>
                                        <SelectItem value="TINGGI">⭐⭐⭐ Tinggi</SelectItem>
                                        <SelectItem value="SANGAT_TINGGI">⭐⭐⭐⭐ Sangat Tinggi</SelectItem>
                                        <SelectItem value="KRITIS">⭐⭐⭐⭐⭐ Kritis</SelectItem>
                                      </SelectContent>
                                    </Select>
                                  ) : (
                                    renderStars(rec.tingkatUrgensi)
                                  )}
                                </TableCell>
                                <TableCell>
                                  {editingRowId === rec.id ? (
                                    <Input
                                      value={editRowForm.pihakTerkait}
                                      onChange={(e) => setEditRowForm({...editRowForm, pihakTerkait: e.target.value})}
                                      className="h-8 text-xs"
                                      placeholder="Pihak terkait..."
                                    />
                                  ) : (
                                    <div className="text-xs">{rec.pihakTerkait || '-'}</div>
                                  )}
                                </TableCell>
                                <TableCell>
                                  {editingRowId === rec.id ? (
                                    <Input
                                      value={editRowForm.organPerusahaan}
                                      onChange={(e) => setEditRowForm({...editRowForm, organPerusahaan: e.target.value})}
                                      className="h-8 text-xs"
                                      placeholder="Organ perusahaan..."
                                    />
                                  ) : (
                                    <div className="text-xs">{rec.organPerusahaan || '-'}</div>
                                  )}
                                </TableCell>
                                <TableCell>
                                  {editingRowId === rec.id ? (
                                    <div className="flex gap-1 justify-center">
                                      <Button
                                        size="sm"
                                        variant="outline"
                                        onClick={() => saveEditedRow(rec.id)}
                                        className="h-7 px-2 border-green-200 text-green-600"
                                      >
                                        <Save className="w-3 h-3" />
                                      </Button>
                                      <Button
                                        size="sm"
                                        variant="outline"
                                        onClick={cancelEditingRow}
                                        className="h-7 px-2 border-gray-200 text-gray-600"
                                      >
                                        <X className="w-3 h-3" />
                                      </Button>
                                    </div>
                                  ) : (
                                    <div className="flex gap-1 justify-center">
                                      <Button
                                        size="sm"
                                        variant="outline"
                                        onClick={() => startEditingRow(rec)}
                                        className="h-7 px-2 border-blue-200 text-blue-600"
                                        data-tour={recIndex === 0 && table.id === yearTables[0]?.id ? "edit-button" : undefined}
                                      >
                                        <Edit className="w-3 h-3" />
                                      </Button>
                                      <Button
                                        size="sm"
                                        variant="outline"
                                        onClick={() => deleteRow(rec.id)}
                                        className="h-7 px-2 border-red-200 text-red-600"
                                        data-tour={recIndex === 0 && table.id === yearTables[0]?.id ? "delete-button" : undefined}
                                      >
                                        <Trash2 className="w-3 h-3" />
                                      </Button>
                                    </div>
                                  )}
                                </TableCell>
                              </TableRow>
                              ))}

                              {/* Inline Add Row for Rekomendasi */}
                              <TableRow className="bg-blue-50/50 border-t-2 border-blue-200">
                                <TableCell className="text-center text-gray-400">
                                  <Plus className="w-4 h-4 text-blue-500 mx-auto" />
                                </TableCell>
                                <TableCell>
                                  <Textarea
                                    value={inlineRekomendasiForm[table.id]?.isi || ''}
                                    onChange={(e) => setInlineRekomendasiForm({
                                      ...inlineRekomendasiForm,
                                      [table.id]: {...inlineRekomendasiForm[table.id], isi: e.target.value}
                                    })}
                                    placeholder="Ketik deskripsi rekomendasi baru..."
                                    className="text-sm min-h-[60px] bg-white"
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter' && e.ctrlKey && inlineRekomendasiForm[table.id]?.isi?.trim()) {
                                        handleAddRekomendasi(table.id);
                                      }
                                    }}
                                  />
                                </TableCell>
                                <TableCell>
                                  <Input
                                    value={inlineRekomendasiForm[table.id]?.aspekAOI || ''}
                                    onChange={(e) => setInlineRekomendasiForm({
                                      ...inlineRekomendasiForm,
                                      [table.id]: {...inlineRekomendasiForm[table.id], aspekAOI: e.target.value}
                                    })}
                                    placeholder="Opsional..."
                                    className="h-8 text-sm bg-white"
                                  />
                                </TableCell>
                                <TableCell>
                                  <Select
                                    value={inlineRekomendasiForm[table.id]?.tingkatUrgensi || 'TIDAK_ADA'}
                                    onValueChange={(v) => setInlineRekomendasiForm({
                                      ...inlineRekomendasiForm,
                                      [table.id]: {...inlineRekomendasiForm[table.id], tingkatUrgensi: v}
                                    })}
                                  >
                                    <SelectTrigger className="h-8 text-sm bg-white">
                                      <SelectValue />
                                    </SelectTrigger>
                                    <SelectContent>
                                      <SelectItem value="TIDAK_ADA">-</SelectItem>
                                      <SelectItem value="RENDAH">⭐ Rendah</SelectItem>
                                      <SelectItem value="SEDANG">⭐⭐ Sedang</SelectItem>
                                      <SelectItem value="TINGGI">⭐⭐⭐ Tinggi</SelectItem>
                                      <SelectItem value="SANGAT_TINGGI">⭐⭐⭐⭐ Sangat Tinggi</SelectItem>
                                      <SelectItem value="KRITIS">⭐⭐⭐⭐⭐ Kritis</SelectItem>
                                    </SelectContent>
                                  </Select>
                                </TableCell>
                                <TableCell>
                                  <Input
                                    value={inlineRekomendasiForm[table.id]?.pihakTerkait || ''}
                                    onChange={(e) => setInlineRekomendasiForm({
                                      ...inlineRekomendasiForm,
                                      [table.id]: {...inlineRekomendasiForm[table.id], pihakTerkait: e.target.value}
                                    })}
                                    placeholder="Opsional..."
                                    className="h-8 text-sm bg-white"
                                  />
                                </TableCell>
                                <TableCell>
                                  <Input
                                    value={inlineRekomendasiForm[table.id]?.organPerusahaan || ''}
                                    onChange={(e) => setInlineRekomendasiForm({
                                      ...inlineRekomendasiForm,
                                      [table.id]: {...inlineRekomendasiForm[table.id], organPerusahaan: e.target.value}
                                    })}
                                    placeholder="Opsional..."
                                    className="h-8 text-sm bg-white"
                                  />
                                </TableCell>
                                <TableCell className="text-center">
                                  <Button
                                    size="sm"
                                    className="bg-blue-600 hover:bg-blue-700 h-8"
                                    onClick={() => handleAddRekomendasi(table.id)}
                                    disabled={!inlineRekomendasiForm[table.id]?.isi?.trim() || isAddingRekomendasi[table.id]}
                                  >
                                    {isAddingRekomendasi[table.id] ? (
                                      <RefreshCw className="w-4 h-4 animate-spin" />
                                    ) : (
                                      <Plus className="w-4 h-4" />
                                    )}
                                  </Button>
                                </TableCell>
                              </TableRow>
                            </TableBody>
                          </Table>
                        </div>
                      </div>
                    </div>

                    {/* Daftar Saran */}
                    <div className="mb-6" data-tour="saran-table">
                      <h4 className="text-lg font-semibold text-blue-900 mb-3">
                        Daftar Saran ({getTableSaran(table.id).length})
                      </h4>
                      <div className="border border-green-200 rounded-lg overflow-hidden bg-white shadow-sm">
                        <div className="overflow-x-auto">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-green-50">
                                <TableHead className="text-green-900 font-semibold w-12 text-center">NO</TableHead>
                                <TableHead className="text-green-900 font-semibold min-w-[300px]">DESKRIPSI</TableHead>
                                <TableHead className="text-green-900 font-semibold w-32">ASPEK AOI</TableHead>
                                <TableHead className="text-green-900 font-semibold w-32 text-center">URGENSI</TableHead>
                                <TableHead className="text-green-900 font-semibold w-40">PIHAK TERKAIT</TableHead>
                                <TableHead className="text-green-900 font-semibold w-40">ORGAN PERUSAHAAN</TableHead>
                                <TableHead className="text-green-900 font-semibold w-32 text-center">AKSI</TableHead>
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {/* Existing Saran rows */}
                              {getTableSaran(table.id).map((rec) => (
                                <TableRow key={rec.id} className="hover:bg-green-50/50 border-b border-green-100">
                                  <TableCell className="font-medium text-center text-green-900">{rec.no}</TableCell>
                                  <TableCell>
                                    {editingRowId === rec.id ? (
                                      <Textarea
                                        value={editRowForm.isi}
                                        onChange={(e) => setEditRowForm({...editRowForm, isi: e.target.value})}
                                        className="text-xs min-h-[60px]"
                                        placeholder="Deskripsi saran..."
                                      />
                                    ) : (
                                      <div className="text-sm leading-relaxed text-gray-800">{rec.isi || '-'}</div>
                                    )}
                                  </TableCell>
                                  <TableCell>
                                    {editingRowId === rec.id ? (
                                      <Input
                                        value={editRowForm.aspekAOI}
                                        onChange={(e) => setEditRowForm({...editRowForm, aspekAOI: e.target.value})}
                                        className="h-8 text-xs"
                                        placeholder="Aspek AOI..."
                                      />
                                    ) : (
                                      <div className="text-xs">{rec.aspekAOI || '-'}</div>
                                    )}
                                  </TableCell>
                                  <TableCell className="text-center">
                                    {editingRowId === rec.id ? (
                                      <Select
                                        value={editRowForm.tingkatUrgensi}
                                        onValueChange={(v) => setEditRowForm({...editRowForm, tingkatUrgensi: v})}
                                      >
                                        <SelectTrigger className="h-8 text-xs">
                                          <SelectValue />
                                        </SelectTrigger>
                                        <SelectContent>
                                          <SelectItem value="TIDAK_ADA">-</SelectItem>
                                          <SelectItem value="RENDAH">⭐ Rendah</SelectItem>
                                          <SelectItem value="SEDANG">⭐⭐ Sedang</SelectItem>
                                          <SelectItem value="TINGGI">⭐⭐⭐ Tinggi</SelectItem>
                                          <SelectItem value="SANGAT_TINGGI">⭐⭐⭐⭐ Sangat Tinggi</SelectItem>
                                          <SelectItem value="KRITIS">⭐⭐⭐⭐⭐ Kritis</SelectItem>
                                        </SelectContent>
                                      </Select>
                                    ) : (
                                      renderStars(rec.tingkatUrgensi)
                                    )}
                                  </TableCell>
                                  <TableCell>
                                    {editingRowId === rec.id ? (
                                      <Input
                                        value={editRowForm.pihakTerkait}
                                        onChange={(e) => setEditRowForm({...editRowForm, pihakTerkait: e.target.value})}
                                        className="h-8 text-xs"
                                        placeholder="Pihak terkait..."
                                      />
                                    ) : (
                                      <div className="text-xs">{rec.pihakTerkait || '-'}</div>
                                    )}
                                  </TableCell>
                                  <TableCell>
                                    {editingRowId === rec.id ? (
                                      <Input
                                        value={editRowForm.organPerusahaan}
                                        onChange={(e) => setEditRowForm({...editRowForm, organPerusahaan: e.target.value})}
                                        className="h-8 text-xs"
                                        placeholder="Organ perusahaan..."
                                      />
                                    ) : (
                                      <div className="text-xs">{rec.organPerusahaan || '-'}</div>
                                    )}
                                  </TableCell>
                                  <TableCell>
                                    {editingRowId === rec.id ? (
                                      <div className="flex gap-1 justify-center">
                                        <Button
                                          size="sm"
                                          variant="outline"
                                          onClick={() => saveEditedRow(rec.id)}
                                          className="h-7 px-2 border-blue-200 text-blue-600"
                                        >
                                          <Save className="w-3 h-3" />
                                        </Button>
                                        <Button
                                          size="sm"
                                          variant="outline"
                                          onClick={cancelEditing}
                                          className="h-7 px-2 border-gray-200 text-gray-600"
                                        >
                                          <X className="w-3 h-3" />
                                        </Button>
                                      </div>
                                    ) : (
                                      <div className="flex gap-1 justify-center">
                                        <Button
                                          size="sm"
                                          variant="outline"
                                          onClick={() => startEditingRow(rec)}
                                          className="h-7 px-2 text-blue-600 hover:text-blue-700"
                                        >
                                          <Edit className="w-3 h-3" />
                                        </Button>
                                        <Button
                                          size="sm"
                                          variant="outline"
                                          onClick={() => handleDeleteRecommendation(rec.id)}
                                          className="h-7 px-2 text-red-600 hover:text-red-700"
                                        >
                                          <Trash2 className="w-3 h-3" />
                                        </Button>
                                      </div>
                                    )}
                                  </TableCell>
                                </TableRow>
                              ))}

                              {/* Inline Add Row for Saran */}
                              <TableRow className="bg-green-50/50 border-t-2 border-green-200">
                                <TableCell className="text-center text-gray-400">
                                  <Plus className="w-4 h-4 text-green-500 mx-auto" />
                                </TableCell>
                                <TableCell>
                                  <Textarea
                                    value={inlineSaranForm[table.id]?.isi || ''}
                                    onChange={(e) => setInlineSaranForm({
                                      ...inlineSaranForm,
                                      [table.id]: {...inlineSaranForm[table.id], isi: e.target.value}
                                    })}
                                    placeholder="Ketik deskripsi saran baru..."
                                    className="text-sm min-h-[60px] bg-white"
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter' && e.ctrlKey && inlineSaranForm[table.id]?.isi?.trim()) {
                                        handleAddSaran(table.id);
                                      }
                                    }}
                                  />
                                </TableCell>
                                <TableCell>
                                  <Input
                                    value={inlineSaranForm[table.id]?.aspekAOI || ''}
                                    onChange={(e) => setInlineSaranForm({
                                      ...inlineSaranForm,
                                      [table.id]: {...inlineSaranForm[table.id], aspekAOI: e.target.value}
                                    })}
                                    placeholder="Opsional..."
                                    className="h-8 text-sm bg-white"
                                  />
                                </TableCell>
                                <TableCell>
                                  <Select
                                    value={inlineSaranForm[table.id]?.tingkatUrgensi || 'TIDAK_ADA'}
                                    onValueChange={(v) => setInlineSaranForm({
                                      ...inlineSaranForm,
                                      [table.id]: {...inlineSaranForm[table.id], tingkatUrgensi: v}
                                    })}
                                  >
                                    <SelectTrigger className="h-8 text-sm bg-white">
                                      <SelectValue />
                                    </SelectTrigger>
                                    <SelectContent>
                                      <SelectItem value="TIDAK_ADA">-</SelectItem>
                                      <SelectItem value="RENDAH">⭐ Rendah</SelectItem>
                                      <SelectItem value="SEDANG">⭐⭐ Sedang</SelectItem>
                                      <SelectItem value="TINGGI">⭐⭐⭐ Tinggi</SelectItem>
                                      <SelectItem value="SANGAT_TINGGI">⭐⭐⭐⭐ Sangat Tinggi</SelectItem>
                                      <SelectItem value="KRITIS">⭐⭐⭐⭐⭐ Kritis</SelectItem>
                                    </SelectContent>
                                  </Select>
                                </TableCell>
                                <TableCell>
                                  <Input
                                    value={inlineSaranForm[table.id]?.pihakTerkait || ''}
                                    onChange={(e) => setInlineSaranForm({
                                      ...inlineSaranForm,
                                      [table.id]: {...inlineSaranForm[table.id], pihakTerkait: e.target.value}
                                    })}
                                    placeholder="Opsional..."
                                    className="h-8 text-sm bg-white"
                                  />
                                </TableCell>
                                <TableCell>
                                  <Input
                                    value={inlineSaranForm[table.id]?.organPerusahaan || ''}
                                    onChange={(e) => setInlineSaranForm({
                                      ...inlineSaranForm,
                                      [table.id]: {...inlineSaranForm[table.id], organPerusahaan: e.target.value}
                                    })}
                                    placeholder="Opsional..."
                                    className="h-8 text-sm bg-white"
                                  />
                                </TableCell>
                                <TableCell className="text-center">
                                  <Button
                                    size="sm"
                                    className="bg-green-600 hover:bg-green-700 h-8"
                                    onClick={() => handleAddSaran(table.id)}
                                    disabled={!inlineSaranForm[table.id]?.isi?.trim() || isAddingSaran[table.id]}
                                  >
                                    {isAddingSaran[table.id] ? (
                                      <RefreshCw className="w-4 h-4 animate-spin" />
                                    ) : (
                                      <Plus className="w-4 h-4" />
                                    )}
                                  </Button>
                                </TableCell>
                              </TableRow>
                            </TableBody>
                          </Table>
                        </div>
                      </div>
                    </div>
                  </CardContent>
                )}
              </Card>
            ))}

            {yearTables.length === 0 && (
              <Card className="border-0 shadow-lg bg-gradient-to-r from-white to-blue-50">
                <CardContent className="p-8">
                  <div className="text-center text-blue-600">
                    <FileText className="h-16 w-16 mx-auto mb-4 text-blue-400" />
                    <h3 className="text-lg font-semibold mb-2">Belum Ada Tabel AOI</h3>
                    <p className="text-sm text-blue-700 mb-4">
                      Buat tabel AOI pertama menggunakan form di atas
                    </p>
                    <Button
                      onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}
                      className="bg-blue-600 hover:bg-blue-700"
                    >
                      <Plus className="w-4 h-4 mr-2" />
                      Isi Form di Atas
                    </Button>
                  </div>
                </CardContent>
              </Card>
            )}
          </div>
        </div>
      </div>

    </>
  );
};

export default AOIManagement;
