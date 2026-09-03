export interface Biro {
  id: string;
  name: string;
  pic?: string;
  desc?: string;
}

export interface Department {
  id: string;
  name: string;
  code?: string;
  icon: 'Wrench' | 'CircleDollarSign' | 'Users' | 'Truck' | 'ShieldCheck' | 'Laptop';
  description?: string;
  biros: Biro[];
}

// Daftar 6 bulan untuk tombol
export const monthList = [
  'Januari',
  'Februari',
  'Maret',
  'April',
  'Mei',
  'Juni'
];

export const departmentsData: Department[] = [
  {
    id: "dept-1",
    name: "Departemen Perencanaan Desain",
    icon: "Wrench",
    biros: [
      { id: "b-101", name: "Biro Dokumen dan Perencanaan" },
      { id: "b-102", name: "Biro Dukungan Logistik Terpadu" },
      { id: "b-103", name: "Biro Dukungan & Administrasi" }
    ]
  },
  {
    id: "dept-2",
    name: "Departemen Desain Dasar",
    icon: "CircleDollarSign",
    biros: [
      { id: "b-201", name: "Biro Desain Pengembangan Desain" },
      { id: "b-202", name: "Biro Desain Dasar Kapal Selam" },
      { id: "b-203", name: "Biro Desain Dasar Non Kapal" },
      { id: "b-204", name: "Biro Desain Dasar Kapal Permukaan" }
    ]
  },
  {
    id: "dept-3",
    name: "Departemen Struktur dan Perlengkapan Lambung",
    icon: "Users",
    biros: [
      { id: "b-301", name: "Biro Desain Struktur Lambung" },
      { id: "b-302", name: "Biro Desain Akomodasi" },
      { id: "b-303", name: "Biro Desain Perlengkapan Lambung" },
      { id: "b-304", name: "Biro Desain Produksi Lambung" }
    ]
  },
  {
    id: "dept-4",
    name: "Departemen Struktur & Perlengkapan Permesinan",
    icon: "Truck",
    biros: [
      { id: "b-401", name: "Biro Sistem Propulsi" },
      { id: "b-402", name: "Biro Pengaturan Permesinan" },
      { id: "b-403", name: "Biro Sistem HVAC dan Permesinan Geladak" }
    ]
  },
  {
    id: "dept-5",
    name: "Departemen Perlengkapan Listrik & Elektronika",
    icon: "ShieldCheck",
    biros: [
      { id: "b-501", name: "Biro Sistem Kelistrikan" },
      { id: "b-502", name: "Biro Sistem Kontrol dan Otomasi" },
      { id: "b-503", name: "Biro Elektronika" }
    ]
  },
  {
    id: "dept-6",
    name: "Departemen HPS",
    icon: "Laptop",
    biros: [
      { id: "b-601", name: "Biro HPS Material" },
      { id: "b-602", name: "Biro HPS Jasa & Investasi" }
    ]
  }
];