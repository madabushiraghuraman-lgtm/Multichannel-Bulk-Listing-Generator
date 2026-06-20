import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileSpreadsheet, 
  Upload, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  ArrowRight, 
  Copy, 
  Check, 
  Plus, 
  Trash2, 
  RefreshCw, 
  Search, 
  BookOpen, 
  ExternalLink,
  Smartphone,
  CheckCircle,
  HelpCircle,
  Grid
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Define TS Interfaces for Master Schema
export interface MasterRow {
  id: string; // Internal unique ID for table keys
  Master_SKU: string;
  Brand: string;
  Style_Code: string;
  Article_Number: string;
  Product_Title: string;
  MRP: number | string;
  Selling_Price: number | string;
  UK_Size: string | number;
  Euro_Size: string | number;
  Color: string;
  Product_Description: string;
  HSN_Code: string;
  Main_Image_URL: string;
  Image_URL_2: string;
  Image_URL_3: string;
}

// 15 Standard Master Columns Requested
const MASTER_COLUMNS: (keyof Omit<MasterRow, 'id'>)[] = [
  'Master_SKU', 'Brand', 'Style_Code', 'Article_Number', 'Product_Title', 
  'MRP', 'Selling_Price', 'UK_Size', 'Euro_Size', 'Color', 
  'Product_Description', 'HSN_Code', 'Main_Image_URL', 'Image_URL_2', 'Image_URL_3'
];

// Helper to normalize strings for robust fuzzy match lookup
const normalizeKey = (key: string): string => {
  return key.trim().toLowerCase().replace(/[\s_\-\(\)\/]/g, '');
};

// Map file columns to master template properties intelligently
const fuzzyColumnMap: Record<string, keyof Omit<MasterRow, 'id'>> = {
  // Master SKU
  mastersku: 'Master_SKU',
  sku: 'Master_SKU',
  sellersku: 'Master_SKU',
  sellerskuid: 'Master_SKU',
  vendorsku: 'Master_SKU',
  vendorskucode: 'Master_SKU',
  productsku: 'Master_SKU',
  
  // Brand
  brand: 'Brand',
  brandname: 'Brand',
  brand_name: 'Brand',
  maker: 'Brand',
  company: 'Brand',

  // Style Code
  stylecode: 'Style_Code',
  styleno: 'Style_Code',
  stylenumber: 'Style_Code',
  style_no: 'Style_Code',

  // Article Number
  articlenumber: 'Article_Number',
  articleno: 'Article_Number',
  article_no: 'Article_Number',
  articlecode: 'Article_Number',
  vendorarticlenumber: 'Article_Number',

  // Product Title
  producttitle: 'Product_Title',
  title: 'Product_Title',
  itemname: 'Product_Title',
  itemnameakatitle: 'Product_Title',
  productname: 'Product_Title',
  displayname: 'Product_Title',
  productdisplayname: 'Product_Title',

  // MRP
  mrp: 'MRP',
  mrp_inr: 'MRP',
  mrpinr: 'MRP',
  maximumretailprice: 'MRP',
  retailprice: 'MRP',

  // Selling Price
  sellingprice: 'Selling_Price',
  selling_price_inr: 'Selling_Price',
  sellingpriceinr: 'Selling_Price',
  price: 'Selling_Price',
  yourprice: 'Selling_Price',
  isp: 'Selling_Price',
  yoursellingpriceinr: 'Selling_Price',

  // UK Size
  uksize: 'UK_Size',
  ukindiasize: 'UK_Size',
  ukindiasizephoto: 'UK_Size',
  shoesize: 'UK_Size',
  sizeuk: 'UK_Size',

  // Euro Size
  eurosize: 'Euro_Size',
  eursize: 'Euro_Size',
  euro_size: 'Euro_Size',
  sizeeuro: 'Euro_Size',

  // Color
  color: 'Color',
  colour: 'Color',
  prominentcolor: 'Color',
  prominentcolour: 'Color',

  // Product Description
  productdescription: 'Product_Description',
  description: 'Product_Description',
  details: 'Product_Description',
  productdetails: 'Product_Description',

  // HSN Code
  hsncode: 'HSN_Code',
  hsn: 'HSN_Code',
  hsn_code: 'HSN_Code',

  // Main Image
  mainimageurl: 'Main_Image_URL',
  mainimage: 'Main_Image_URL',
  image1: 'Main_Image_URL',
  imageurl1: 'Main_Image_URL',
  frontimage: 'Main_Image_URL',

  // Image 2
  imageurl2: 'Image_URL_2',
  image2: 'Image_URL_2',
  otherimageurl1: 'Image_URL_2',
  sideimage: 'Image_URL_2',

  // Image 3
  imageurl3: 'Image_URL_3',
  image3: 'Image_URL_3',
  otherimageurl2: 'Image_URL_3',
  backimage: 'Image_URL_3',
};

// Premium high-fidelity sample footwear data to allow immediate live playing
const SAMPLE_FOOTWEAR_DATA: MasterRow[] = [
  {
    id: 'sample-1',
    Master_SKU: 'AETHER-RUN-RD-08',
    Brand: 'Aether',
    Style_Code: 'AE-RUN-22',
    Article_Number: 'AETH-9011',
    Product_Title: 'Aether Air-Step Pro Running Shoes - Crimson Flame Red',
    MRP: 4999,
    Selling_Price: 3499,
    UK_Size: 8,
    Euro_Size: 42,
    Color: 'Crimson Flame Red',
    Product_Description: 'Engineered mesh upper for high breathability, paired with ultra-light reactive foam soles.',
    HSN_Code: '64041190',
    Main_Image_URL: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?auto=format&fit=crop&w=600&q=80',
    Image_URL_2: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?auto=format&fit=crop&w=600&q=80',
    Image_URL_3: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?auto=format&fit=crop&w=600&q=80'
  },
  {
    id: 'sample-2',
    Master_SKU: 'AETHER-RUN-RD-09',
    Brand: 'Aether',
    Style_Code: 'AE-RUN-22',
    Article_Number: 'AETH-9011',
    Product_Title: 'Aether Air-Step Pro Running Shoes - Crimson Flame Red',
    MRP: 4999,
    Selling_Price: 3499,
    UK_Size: 9,
    Euro_Size: 43,
    Color: 'Crimson Flame Red',
    Product_Description: 'Engineered mesh upper for high breathability, paired with ultra-light reactive foam soles.',
    HSN_Code: '64041190',
    Main_Image_URL: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?auto=format&fit=crop&w=600&q=80',
    Image_URL_2: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?auto=format&fit=crop&w=600&q=80',
    Image_URL_3: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?auto=format&fit=crop&w=600&q=80'
  },
  {
    id: 'sample-3',
    Master_SKU: 'VG-BRN-BOOT-10',
    Brand: 'Vanguard Boots',
    Style_Code: 'VG-LTHR-05',
    Article_Number: 'VG-7762',
    Product_Title: 'Vanguard Handcrafted Artisan Leather Boots - Espresso Brown',
    MRP: 8999,
    Selling_Price: 6499,
    UK_Size: 10,
    Euro_Size: 44,
    Color: 'Espresso Brown',
    Product_Description: 'Individually constructed with premium full-grain Italian calf leather. Water-resistant welt stitching.',
    HSN_Code: '64035111',
    Main_Image_URL: 'https://images.unsplash.com/photo-1608256246200-53e635b5b65f?auto=format&fit=crop&w=600&q=80',
    Image_URL_2: 'https://images.unsplash.com/photo-1608256246200-53e635b5b65f?auto=format&fit=crop&w=600&q=80',
    Image_URL_3: 'https://images.unsplash.com/photo-1608256246200-53e635b5b65f?auto=format&fit=crop&w=600&q=80'
  },
  {
    id: 'sample-4',
    Master_SKU: 'CLD-FOAM-BK-07',
    Brand: 'CloudWalkers',
    Style_Code: 'CW-SLIDE-01',
    Article_Number: 'CW-1140',
    Product_Title: 'CloudWalkers Memory Foam Summer Slides - Obsidian Black',
    MRP: 1999,
    Selling_Price: 1399,
    UK_Size: 7,
    Euro_Size: 41,
    Color: 'Obsidian Black',
    Product_Description: 'Plush dual-density memory foam footbed with orthopedic arch support and textured water-grip sole.',
    HSN_Code: '64029990',
    Main_Image_URL: 'https://images.unsplash.com/photo-1603808033192-082d6919d3e1?auto=format&fit=crop&w=600&q=80',
    Image_URL_2: 'https://images.unsplash.com/photo-1603808033192-082d6919d3e1?auto=format&fit=crop&w=600&q=80',
    Image_URL_3: 'https://images.unsplash.com/photo-1603808033192-082d6919d3e1?auto=format&fit=crop&w=600&q=80'
  }
];

export default function App() {
  // State variables
  const [activeTab, setActiveTab] = useState<'workbench' | 'portability' | 'reference'>('workbench');
  const [products, setProducts] = useState<MasterRow[]>(SAMPLE_FOOTWEAR_DATA);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [editingCell, setEditingCell] = useState<{ id: string; col: keyof MasterRow } | null>(null);
  const [editValue, setEditValue] = useState<string>('');
  
  // File upload state info
  const [fileName, setFileName] = useState<string>('Sample Dataset pre-loaded');
  const [fileHeaders, setFileHeaders] = useState<string[]>(MASTER_COLUMNS);
  const [uploadFeedback, setUploadFeedback] = useState<{ type: 'success' | 'warning' | null; message: string }>({
    type: 'success',
    message: 'Playground successfully loaded with curated sample footwear listings.'
  });

  // State to check code copied
  const [copied, setCopied] = useState<boolean>(false);

  // Pagination parameters
  const rowsPerPage = 5;
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Column validation - counts columns mapped correctly from Master templates
  const mappedStatus = useMemo(() => {
    const status: Record<string, boolean> = {};
    MASTER_COLUMNS.forEach(col => {
      // If our current fileHeaders matches or has fuzzy equivalent mapped
      status[col] = fileHeaders.some(fh => {
        const normFH = normalizeKey(fh);
        const mappedCol = fuzzyColumnMap[normFH];
        return mappedCol === col || fh === col;
      });
    });
    return status;
  }, [fileHeaders]);

  const numColumnsMapped = useMemo(() => {
    return Object.values(mappedStatus).filter(b => b === true).length;
  }, [mappedStatus]);

  // Handle uploaded Excel sheet
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Convert sheet to json
        const rawJson = XLSX.utils.sheet_to_json<any>(sheet);
        if (rawJson.length === 0) {
          setUploadFeedback({
            type: 'warning',
            message: 'Target spreadsheet is empty or has no recognizable rows.'
          });
          return;
        }

        // Detect all headers in the uploaded sheets
        const detectedHeaders: string[] = [];
        const firstRow = rawJson[0];
        Object.keys(firstRow).forEach(k => detectedHeaders.push(k));
        setFileHeaders(detectedHeaders);

        // Convert raw json keys using fuzzy mapper
        const mappedRows: MasterRow[] = rawJson.map((row, index) => {
          const newRow: any = { id: `uploaded-${Date.now()}-${index}` };
          
          // Pre-populate empty properties
          MASTER_COLUMNS.forEach(col => {
            newRow[col] = '';
          });

          // Match each key in the row to our schema
          Object.keys(row).forEach(rowKey => {
            const rawVal = row[rowKey];
            const normalizedRowKey = normalizeKey(rowKey);
            const masterKey = fuzzyColumnMap[normalizedRowKey];
            
            if (masterKey) {
              newRow[masterKey] = rawVal;
            }
          });

          return newRow as MasterRow;
        });

        setProducts(mappedRows);
        setCurrentPage(1);

        // Calculate count of mapped columns
        const successfullyMappedCount = MASTER_COLUMNS.filter(col => {
          return detectedHeaders.some(fh => {
            const normFH = normalizeKey(fh);
            const mappedCol = fuzzyColumnMap[normFH];
            return mappedCol === col || fh === col;
          });
        }).length;

        if (successfullyMappedCount >= 10) {
          setUploadFeedback({
            type: 'success',
            message: `Excel processed successfully! Detected and mapped ${successfullyMappedCount}/15 standard footwear columns.`
          });
        } else {
          setUploadFeedback({
            type: 'warning',
            message: `Partially processed! Cleaned standard headers, but mapped only ${successfullyMappedCount}/15 columns. (Missing parameters can be edited or manually typed below)`
          });
        }
      } catch (err) {
        console.error(err);
        setUploadFeedback({
          type: 'warning',
          message: 'Error reading file. Ensure it is a valid, uncorrupted Excel (.xlsx) file.'
        });
      }
    };
    reader.readAsBinaryString(file);
  };

  // Trigger manual workbook upload click
  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  // Master Sample download functionality so users have a beautiful reference template
  const downloadMasterSample = () => {
    const rawSampleData = SAMPLE_FOOTWEAR_DATA.map(({ id, ...rest }) => rest);
    const ws = XLSX.utils.json_to_sheet(rawSampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Master Footwear Registry");
    XLSX.writeFile(wb, "Footwear_Master_Sample.xlsx");
  };

  // Add a blank row
  const handleAddRow = () => {
    const freshRow: MasterRow = {
      id: `manual-${Date.now()}`,
      Master_SKU: `NEW-SKU-${Math.floor(1000 + Math.random() * 9000)}`,
      Brand: 'Brand Name',
      Style_Code: 'SC-100',
      Article_Number: 'ART-200',
      Product_Title: 'Product Model Title',
      MRP: 3999,
      Selling_Price: 2999,
      UK_Size: 8,
      Euro_Size: 42,
      Color: 'Color shade',
      Product_Description: 'Standard description of the master footwear.',
      HSN_Code: '64041190',
      Main_Image_URL: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff',
      Image_URL_2: '',
      Image_URL_3: ''
    };
    setProducts([freshRow, ...products]);
    setCurrentPage(1);
  };

  // Delete product row
  const handleDeleteRow = (id: string) => {
    const cleaned = products.filter(p => p.id !== id);
    setProducts(cleaned);
    // Adjust page bounds if necessary
    const maxPage = Math.max(1, Math.ceil(cleaned.length / rowsPerPage));
    if (currentPage > maxPage) {
      setCurrentPage(maxPage);
    }
  };

  // Cell editing logic
  const startEditing = (id: string, col: keyof MasterRow, currentVal: any) => {
    setEditingCell({ id, col });
    setEditValue(String(currentVal ?? ''));
  };

  const saveCellEdit = () => {
    if (!editingCell) return;
    const { id, col } = editingCell;
    setProducts(products.map(p => {
      if (p.id === id) {
        return {
          ...p,
          [col]: isNaN(Number(editValue)) || editValue.trim() === '' ? editValue : Number(editValue)
        };
      }
      return p;
    }));
    setEditingCell(null);
  };

  // Reset to sample data
  const resetToSample = () => {
    setProducts(SAMPLE_FOOTWEAR_DATA);
    setFileName('Sample Dataset restored');
    setFileHeaders(MASTER_COLUMNS);
    setCurrentPage(1);
    setUploadFeedback({
      type: 'success',
      message: 'Restored original interactive template database containing footwear examples.'
    });
  };

  // Search filter
  const filteredProducts = useMemo(() => {
    if (!searchQuery.trim()) return products;
    const q = searchQuery.toLowerCase();
    return products.filter(p => 
      p.Master_SKU.toLowerCase().includes(q) ||
      p.Brand.toLowerCase().includes(q) ||
      p.Product_Title.toLowerCase().includes(q) ||
      p.Color.toLowerCase().includes(q) ||
      String(p.UK_Size).toLowerCase().includes(q)
    );
  }, [products, searchQuery]);

  // Paginated list
  const paginatedProducts = useMemo(() => {
    const offset = (currentPage - 1) * rowsPerPage;
    return filteredProducts.slice(offset, offset + rowsPerPage);
  }, [filteredProducts, currentPage]);

  const totalPages = Math.max(1, Math.ceil(filteredProducts.length / rowsPerPage));

  // --- Marketplace Converters & Downloads ---

  // Amazon Mapping & Export
  const downloadAmazon = () => {
    const mapped = products.map(prod => ({
      'Seller SKU': prod.Master_SKU,
      'Brand Name': prod.Brand,
      'Model number': prod.Style_Code || prod.Article_Number,
      'Item Name (aka Title)': prod.Product_Title,
      'Product Description': prod.Product_Description,
      'Maximum Retail Price': prod.MRP,
      'Your price': prod.Selling_Price,
      'Footwear Size System': 'UK',
      'Shoe Size': prod.UK_Size,
      'Color': prod.Color,
      'Main Image URL': prod.Main_Image_URL,
      'HSN Code': prod.HSN_Code
    }));
    
    const ws = XLSX.utils.json_to_sheet(mapped);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Amazon Upload Template");
    XLSX.writeFile(wb, "Amazon_Bulk_Upload.xlsx");
  };

  // Flipkart Mapping & Export
  const downloadFlipkart = () => {
    const mapped = products.map(prod => ({
      'Seller SKU ID': prod.Master_SKU,
      'Brand': prod.Brand,
      'Style Code': prod.Style_Code,
      'Article Number': prod.Article_Number,
      'MRP (INR)': prod.MRP,
      'Your selling price (INR)': prod.Selling_Price,
      'UK/India Size': prod.UK_Size,
      'Euro Size': prod.Euro_Size,
      'Color': prod.Color,
      'Description': prod.Product_Description,
      'Main Image URL': prod.Main_Image_URL,
      'Other Image URL 1': prod.Image_URL_2,
      'Other Image URL 2': prod.Image_URL_3,
      'HSN': prod.HSN_Code
    }));

    const ws = XLSX.utils.json_to_sheet(mapped);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Flipkart Upload Template");
    XLSX.writeFile(wb, "Flipkart_Bulk_Upload.xlsx");
  };

  // Myntra Mapping & Export
  const downloadMyntra = () => {
    const mapped = products.map(prod => ({
      'vendorSkuCode': prod.Master_SKU,
      'brand': prod.Brand,
      'vendorArticleNumber': prod.Article_Number,
      'productDisplayName': prod.Product_Title,
      'MRP': prod.MRP,
      'ISP': prod.Selling_Price,
      'UK Size': prod.UK_Size,
      'EURO Size': prod.Euro_Size,
      'Prominent Colour': prod.Color,
      'Product Details': prod.Product_Description,
      'Front Image': prod.Main_Image_URL,
      'Side Image': prod.Image_URL_2,
      'Back Image': prod.Image_URL_3,
      'HSN': prod.HSN_Code
    }));

    const ws = XLSX.utils.json_to_sheet(mapped);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Myntra Upload Template");
    XLSX.writeFile(wb, "Myntra_Bulk_Upload.xlsx");
  };

  // Full mappings execution
  const handleBulkExportAll = () => {
    downloadAmazon();
    downloadFlipkart();
    downloadMyntra();
  };

  // Python Streamlit App Script Code
  const streamLitCodeStr = `import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Footwear Multichannel Bulk Listing Generator",
    page_icon="👟",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .reportview-container .main .block-container { max-width: 1200px; }
    .stButton>button { width: 100%; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)

# Title
st.title("👟 Footwear Multichannel Bulk Listing Generator")
st.markdown("""
This Python Streamlit utility processes your standard master footwear registries (.xlsx)
and instantly formats them into marketplace-sanctioned templates for **Amazon**, **Flipkart**, and **Myntra**.
""")

# File Uploader
uploaded_file = st.file_uploader("Upload Master Excel Sheet (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read uploaded sheet
        df_master = pd.read_excel(uploaded_file)
        
        st.subheader("📋 Master Data Preview")
        st.dataframe(df_master.head(8), use_container_width=True)
        st.success(f"Successfully processed {len(df_master)} products with {len(df_master.columns)} attributes.")

        # Mapping templates operation
        st.subheader("🚀 Transform & Output Templates")
        
        # 1. Amazon Output Template Setup
        amazon_df = pd.DataFrame()
        amazon_df['Seller SKU'] = df_master.get('Master_SKU', pd.Series(dtype='object'))
        amazon_df['Brand Name'] = df_master.get('Brand', pd.Series(dtype='object'))
        amazon_df['Model number'] = df_master.get('Style_Code', df_master.get('Article_Number', pd.Series(dtype='object')))
        amazon_df['Item Name (aka Title)'] = df_master.get('Product_Title', pd.Series(dtype='object'))
        amazon_df['Product Description'] = df_master.get('Product_Description', pd.Series(dtype='object'))
        amazon_df['Maximum Retail Price'] = df_master.get('MRP', pd.Series(dtype='float64'))
        amazon_df['Your price'] = df_master.get('Selling_Price', pd.Series(dtype='float64'))
        amazon_df['Footwear Size System'] = "UK"
        amazon_df['Shoe Size'] = df_master.get('UK_Size', pd.Series(dtype='object'))
        amazon_df['Color'] = df_master.get('Color', pd.Series(dtype='object'))
        amazon_df['Main Image URL'] = df_master.get('Main_Image_URL', pd.Series(dtype='object'))
        amazon_df['HSN Code'] = df_master.get('HSN_Code', pd.Series(dtype='object'))

        # 2. Flipkart Output Template Setup
        flipkart_df = pd.DataFrame()
        flipkart_df['Seller SKU ID'] = df_master.get('Master_SKU', pd.Series(dtype='object'))
        flipkart_df['Brand'] = df_master.get('Brand', pd.Series(dtype='object'))
        flipkart_df['Style Code'] = df_master.get('Style_Code', pd.Series(dtype='object'))
        flipkart_df['Article Number'] = df_master.get('Article_Number', pd.Series(dtype='object'))
        flipkart_df['MRP (INR)'] = df_master.get('MRP', pd.Series(dtype='float64'))
        flipkart_df['Your selling price (INR)'] = df_master.get('Selling_Price', pd.Series(dtype='float64'))
        flipkart_df['UK/India Size'] = df_master.get('UK_Size', pd.Series(dtype='object'))
        flipkart_df['Euro Size'] = df_master.get('Euro_Size', pd.Series(dtype='object'))
        flipkart_df['Color'] = df_master.get('Color', pd.Series(dtype='object'))
        flipkart_df['Description'] = df_master.get('Product_Description', pd.Series(dtype='object'))
        flipkart_df['Main Image URL'] = df_master.get('Main_Image_URL', pd.Series(dtype='object'))
        flipkart_df['Other Image URL 1'] = df_master.get('Image_URL_2', pd.Series(dtype='object'))
        flipkart_df['Other Image URL 2'] = df_master.get('Image_URL_3', pd.Series(dtype='object'))
        flipkart_df['HSN'] = df_master.get('HSN_Code', pd.Series(dtype='object'))

        # 3. Myntra Output Template Setup
        myntra_df = pd.DataFrame()
        myntra_df['vendorSkuCode'] = df_master.get('Master_SKU', pd.Series(dtype='object'))
        myntra_df['brand'] = df_master.get('Brand', pd.Series(dtype='object'))
        myntra_df['vendorArticleNumber'] = df_master.get('Article_Number', pd.Series(dtype='object'))
        myntra_df['productDisplayName'] = df_master.get('Product_Title', pd.Series(dtype='object'))
        myntra_df['MRP'] = df_master.get('MRP', pd.Series(dtype='float64'))
        myntra_df['ISP'] = df_master.get('Selling_Price', pd.Series(dtype='float64'))
        myntra_df['UK Size'] = df_master.get('UK_Size', pd.Series(dtype='object'))
        myntra_df['EURO Size'] = df_master.get('Euro_Size', pd.Series(dtype='object'))
        myntra_df['Prominent Colour'] = df_master.get('Color', pd.Series(dtype='object'))
        myntra_df['Product Details'] = df_master.get('Product_Description', pd.Series(dtype='object'))
        myntra_df['Front Image'] = df_master.get('Main_Image_URL', pd.Series(dtype='object'))
        myntra_df['Side Image'] = df_master.get('Image_URL_2', pd.Series(dtype='object'))
        myntra_df['Back Image'] = df_master.get('Image_URL_3', pd.Series(dtype='object'))
        myntra_df['HSN'] = df_master.get('HSN_Code', pd.Series(dtype='object'))

        # Download interface columns
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("### 🅰️ Amazon")
            buffer_amz = io.BytesIO()
            with pd.ExcelWriter(buffer_amz, engine='openpyxl') as writer:
                amazon_df.to_excel(writer, index=False)
            st.download_button(
                label="📥 Amazon Bulk Sheet",
                data=buffer_amz.getvalue(),
                file_name="Amazon_Bulk_Upload.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            st.markdown("### 🅵 Flipkart")
            buffer_fk = io.BytesIO()
            with pd.ExcelWriter(buffer_fk, engine='openpyxl') as writer:
                flipkart_df.to_excel(writer, index=False)
            st.download_button(
                label="📥 Flipkart Bulk Sheet",
                data=buffer_fk.getvalue(),
                file_name="Flipkart_Bulk_Upload.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col3:
            st.markdown("### 🅼 Myntra")
            buffer_myn = io.BytesIO()
            with pd.ExcelWriter(buffer_myn, engine='openpyxl') as writer:
                myntra_df.to_excel(writer, index=False)
            st.download_button(
                label="📥 Myntra Bulk Sheet",
                data=buffer_myn.getvalue(),
                file_name="Myntra_Bulk_Upload.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.write("---")
        st.subheader("💡 Fully Processed Previews")
        tabs = st.tabs(["Amazon Structure", "Flipkart Structure", "Myntra Structure"])
        
        with tabs[0]:
            st.dataframe(amazon_df.head(5), use_container_width=True)
        with tabs[1]:
            st.dataframe(flipkart_df.head(5), use_container_width=True)
        with tabs[2]:
            st.dataframe(myntra_df.head(5), use_container_width=True)

    except Exception as e:
        st.error(f"Error converting data registry: {e}")
else:
    st.info("💡 Please upload an Excel sheet containing standard footwear master registry columns to begin modeling templates.")`;

  const copyToClipboard = () => {
    navigator.clipboard.writeText(streamLitCodeStr);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const downloadPythonScript = () => {
    const blob = new Blob([streamLitCodeStr], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = "app.py";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const downloadBatLauncher = () => {
    const batContent = `@echo off
:: FOOTWEAR BULK LISTING GENERATOR - OFFLINE AUTO-LAUNCHER
:: Designed for madabushiraghuraman's Footwear Registry offline processing.

title Footwear Bulk Listing Generator
color 0b
echo =======================================================================
echo          FOOTWEAR MULTICHANNEL BULK LISTING GENERATOR
echo                    [Windows Offline Launcher]
echo =======================================================================
echo.
echo [1/3] Verifying Python Environment...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo -------------------------------------------------------------------
    echo [ERROR] Python was not found on this computer.
    echo -------------------------------------------------------------------
    echo  To use this offline app, your employee needs Python installed first.
    echo.
    echo  INSTRUCTIONS FOR THE EMPLOYEE:
    echo  1. Please download Python 3.10 or higher from:
    echo     https://www.python.org/downloads/
    echo  2. IMPORTANT: When installing, check the checkbox at the bottom:
    echo     "[X] Add Python.exe to PATH"
    echo  3. After Python is installed, close this window and double-click 
    echo     this script again!
    echo.
    echo Press any key to exit...
    pause >nul
    exit /b
)

echo [OK] Python detected successfully.
echo.
echo [3/3] Installing and verifying required libraries...
echo (This may take a moment on the first startup...)
echo.
pip install streamlit pandas openpyxl
if %errorlevel% neq 0 (
    echo.
    echo [WARNING] Troubleshooting automated library setup...
    python -m pip install --upgrade pip
    python -m pip install streamlit pandas openpyxl
)
echo.
echo [OK] Libraries are fully ready!
echo.
echo [3/3] Launching local Streamlit Server...
echo.
echo =======================================================================
echo  App is starting! A web browser window will automatically open.
echo  Keep this window open while using the generator program.
echo =======================================================================
echo.
streamlit run app.py
pause`;
    const blob = new Blob([batContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = "Launch_Footwear_App.bat";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const downloadReadmeInstructions = () => {
    const readmeContent = `===============================================================================
  INSTRUCTIONS FOR EMPLOYEE: RUNNING THE FOOTWEAR LISTING GENERATOR OFFLINE
===============================================================================

Hi! 

This offline package allows you to process footwear template spreadsheets locally on your computer without needing an active internet connection.

Follow these simple steps to run the application:

STEP 1: INSTALL PYTHON (Required only once)
-------------------------------------------
1. Check if you have Python installed. If not, download Python from this link:
   https://www.python.org/downloads/
2. Double-click the downloaded Python installer.
3. CRITICAL STEP: At the very bottom of the installation window, check the box that says:
   "Add Python.exe to PATH" (or "Add Python to environment variables").
4. Complete the installation wizard.

STEP 2: PREPARE YOUR OFFICE WORKSPACE FOLDER
--------------------------------------------
1. Create a new folder on your Desktop called "Footwear App".
2. Save both of these files into that folder:
   - "app.py" (The core program application script)
   - "Launch_Footwear_App.bat" (The automatic Windows startup launcher)

STEP 3: RUN THE APPLICATION
---------------------------
1. Double-click the "Launch_Footwear_App.bat" file.
2. A command window will open and automatically check/install Streamlit. 
3. After a few seconds, your default web browser (Chrome or Edge) will open automatically to http://localhost:8501.
4. You will see the visual Footwear Multichannel Bulk Listing Generator interface!
5. Upload your master Excel registry files and export your Amazon, Flipkart, or Myntra bulk templates easily.

TIPS:
-----
- Keep the black background terminal window open while using the application. You can minimize it. When you are done, simply close the terminal window to quit.
- To use the app tomorrow, simply double-click the "Launch_Footwear_App.bat" file again!

Developed using Google AI Studio. Perfect for offline warehouse operations!`;
    const blob = new Blob([readmeContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = "Employee_Handover_Instructions.txt";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div id="bulk-listing-app" className="min-h-screen bg-slate-50 flex flex-col font-sans transition-colors duration-300">
      
      {/* HEADER SECTION */}
      <header className="border-b border-slate-200 bg-white shadow-xs sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-blue-600 rounded-xl text-white shadow-md shadow-blue-500/10">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
                Footwear Multichannel Bulk Listing Generator
                <span className="text-xs bg-slate-100 text-slate-600 border border-slate-200 px-2 py-0.5 rounded-full font-mono">v1.1</span>
              </h1>
              <p className="text-xs text-slate-500">Transform master warehouse stock excel sheets to marketplace bulk templates</p>
            </div>
          </div>

          <div className="flex items-center gap-2 self-start md:self-auto bg-slate-100 p-1 rounded-lg border border-slate-200">
            <button
              onClick={() => setActiveTab('workbench')}
              className={`px-4 py-1.5 rounded-md text-xs font-semibold transition-all duration-150 flex items-center gap-1.5 ${
                activeTab === 'workbench'
                  ? 'bg-white text-slate-900 shadow-xs border border-slate-200'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Grid className="w-3.5 h-3.5" />
              Interactive Workbench
            </button>
            <button
              onClick={() => setActiveTab('portability')}
              className={`px-4 py-1.5 rounded-md text-xs font-semibold transition-all duration-150 flex items-center gap-1.5 ${
                activeTab === 'portability'
                  ? 'bg-white text-slate-900 shadow-xs border border-slate-200'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Smartphone className="w-3.5 h-3.5" />
              Streamlit Script (Python)
            </button>
            <button
              onClick={() => setActiveTab('reference')}
              className={`px-4 py-1.5 rounded-md text-xs font-semibold transition-all duration-150 flex items-center gap-1.5 ${
                activeTab === 'reference'
                  ? 'bg-white text-slate-900 shadow-xs border border-slate-200'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <BookOpen className="w-3.5 h-3.5" />
              Mapping Specs
            </button>
          </div>
        </div>
      </header>

      {/* CORE WORKSPACE CONTEXT */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        {activeTab === 'workbench' && (
          <div className="space-y-8">
            
            {/* FILE DROPZONE AND LOAD REFERENCE MOCK ACTIONS */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col justify-between">
                <div>
                  <h2 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-2">📥 Upload Master Inventory Datasheet</h2>
                  <p className="text-xs text-slate-500 mb-4">
                    Upload your standard master footwear listing file in Microsoft Excel format (.xlsx) to match column tags.
                  </p>
                  
                  {/* Real-time feedback alerts */}
                  {uploadFeedback.type && (
                    <div className={`mb-4 p-3 rounded-lg border text-xs flex items-start gap-2.5 ${
                      uploadFeedback.type === 'success' 
                        ? 'bg-emerald-50 border-emerald-200 text-emerald-800' 
                        : 'bg-amber-50 border-amber-200 text-amber-800'
                    }`}>
                      {uploadFeedback.type === 'success' ? (
                        <CheckCircle2 className="w-4.5 h-4.5 text-emerald-600 shrink-0 mt-0.5" />
                      ) : (
                        <AlertCircle className="w-4.5 h-4.5 text-amber-600 shrink-0 mt-0.5" />
                      )}
                      <div>
                        <span className="font-semibold block">{uploadFeedback.type === 'success' ? 'Ready to Convert' : 'Formatting Warning'}</span>
                        {uploadFeedback.message}
                      </div>
                    </div>
                  )}
                </div>

                <div 
                  onClick={triggerFileSelect} 
                  className="border-2 border-dashed border-slate-300 hover:border-blue-500 hover:bg-blue-50/20 active:bg-blue-50/50 transition-all rounded-xl p-8 text-center cursor-pointer group flex flex-col items-center justify-center gap-3 relative overflow-hidden"
                >
                  <input 
                    type="file" 
                    id="excel-uploader"
                    ref={fileInputRef} 
                    onChange={handleFileUpload} 
                    accept=".xlsx" 
                    className="hidden" 
                  />
                  <div className="p-3 bg-slate-100 rounded-full text-slate-600 group-hover:bg-blue-100 group-hover:text-blue-600 transition-all duration-150">
                    <Upload className="w-6 h-6" />
                  </div>
                  <div>
                    <span id="file-helper-label" className="text-sm font-semibold text-slate-700 block">
                      {fileName.length > 32 ? fileName.substring(0, 30) + "..." : fileName}
                    </span>
                    <p className="text-xs text-slate-400 mt-1">Drag and drop or search desktop directories (.xlsx only)</p>
                  </div>
                </div>

                <div className="flex flex-wrap items-center justify-between gap-3 mt-4 pt-4 border-t border-slate-100">
                  <div className="flex items-center gap-2">
                    <button 
                      onClick={resetToSample}
                      className="text-xs text-slate-600 hover:text-blue-600 transition-colors font-medium flex items-center gap-1.5"
                      title="Restore original products placeholder demo"
                    >
                      <RefreshCw className="w-3.5 h-3.5" />
                      Reload Curated Sample
                    </button>
                    <span className="text-slate-300 text-sm">|</span>
                    <span className="text-slate-500 font-mono text-[11px]">Rows loaded: {products.length}</span>
                  </div>

                  <button 
                    onClick={downloadMasterSample}
                    className="text-xs text-blue-600 hover:text-blue-800 transition-colors font-semibold flex items-center gap-1"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Download Sample Excel File
                  </button>
                </div>
              </div>

              {/* LIVE SCHEMA MAPPING INTELLIGENCE BADGES */}
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <h2 className="text-sm font-bold text-slate-800 uppercase tracking-wider">📐 Column Mapper Tags</h2>
                    <span className="text-xs bg-slate-100 text-slate-700 font-mono font-medium px-2 py-0.5 rounded-md border border-slate-200">
                      {numColumnsMapped}/15 Mapped
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-4">
                    The generator scans your headers for synonyms. Below is the active state of key mappings:
                  </p>
                </div>

                <div className="grid grid-cols-2 gap-2 max-h-[220px] overflow-y-auto pr-1">
                  {MASTER_COLUMNS.map((col) => {
                    const mapped = mappedStatus[col];
                    return (
                      <div 
                        key={col} 
                        className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border text-xs transition-colors duration-150 ${
                          mapped 
                            ? 'bg-emerald-50/50 border-emerald-100 text-emerald-800 font-medium' 
                            : 'bg-slate-50 border-slate-200 text-slate-400'
                        }`}
                        title={mapped ? `"${col}" discovered and parsed!` : `"${col}" missing from uploaded headers`}
                      >
                        {mapped ? (
                          <CheckCircle className="w-3.5 h-3.5 text-emerald-600" />
                        ) : (
                          <span className="w-1.5 h-1.5 rounded-full bg-slate-300 ml-1 block" />
                        )}
                        <span className="truncate">{col}</span>
                      </div>
                    );
                  })}
                </div>

                <div className="mt-4 pt-3 border-t border-slate-100">
                  <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-150 text-[11px] text-slate-500">
                    💡 <strong className="text-slate-700">Pro tip:</strong> Double click on any field in the grid below to rewrite, fix, or populate empty cells directly inside the browser!
                  </div>
                </div>
              </div>

            </div>

            {/* INTERACTIVE WORKBENCH BULK EDITING SHEET */}
            <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/70 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                <div className="flex items-center gap-3">
                  <div className="p-1.5 bg-slate-200 text-slate-700 rounded-lg">
                    <Grid className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="text-base font-bold text-slate-850">Direct Spreadsheet Workbench</h3>
                    <p className="text-xs text-slate-500">Search, alter values, add rows, or delete records to adjust templates in real-time</p>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <div className="relative">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input 
                      type="text" 
                      placeholder="Fuzzy search listing..."
                      value={searchQuery}
                      onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }}
                      className="pl-9 pr-4 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-medium focus:outline-hidden focus:border-blue-500 focus:ring-1 focus:ring-blue-500 placeholder-slate-400 w-48 transition-all"
                    />
                  </div>

                  <button
                    onClick={handleAddRow}
                    className="bg-slate-900 hover:bg-slate-800 text-white font-semibold text-xs px-3.5 py-1.5 rounded-lg transition-all duration-150 flex items-center gap-1.5 shadow-sm"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    Add Footwear Row
                  </button>
                </div>
              </div>

              {/* GRID TABLE CONTAINER */}
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs select-none">
                  <thead>
                    <tr className="bg-slate-100 text-slate-600 font-mono uppercase text-[10px] tracking-wider border-b border-slate-200">
                      <th className="py-3 px-4 text-center font-bold">Action</th>
                      <th className="py-3 px-4 font-bold border-l border-slate-250">Master_SKU</th>
                      <th className="py-3 px-4 font-bold">Brand</th>
                      <th className="py-3 px-4 font-bold">Style_Code</th>
                      <th className="py-3 px-4 font-bold">Article_Number</th>
                      <th className="py-3 px-4 font-bold">Product_Title</th>
                      <th className="py-3 px-4 font-bold text-right">MRP</th>
                      <th className="py-3 px-4 font-bold text-right">Selling_Price</th>
                      <th className="py-3 px-4 font-bold text-center">UK_Size</th>
                      <th className="py-3 px-4 font-bold text-center">Euro_Size</th>
                      <th className="py-3 px-4 font-bold">Color</th>
                      <th className="py-3 px-3 min-w-[200px]">Product_Description</th>
                      <th className="py-3 px-4 font-bold text-center">HSN</th>
                      <th className="py-3 px-2">Image URLs</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-sans text-slate-700">
                    {paginatedProducts.map((prod) => (
                      <tr key={prod.id} className="hover:bg-slate-50/80 group duration-100">
                        {/* Remove Action */}
                        <td className="py-2.5 px-4 text-center">
                          <button
                            onClick={() => handleDeleteRow(prod.id)}
                            className="p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded"
                            title="Remove product"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>

                        {/* Editable inputs */}
                        {[
                          { key: 'Master_SKU', align: 'left', bold: true },
                          { key: 'Brand', align: 'left' },
                          { key: 'Style_Code', align: 'left' },
                          { key: 'Article_Number', align: 'left' },
                          { key: 'Product_Title', align: 'left', truncate: true },
                          { key: 'MRP', align: 'right', isNumber: true },
                          { key: 'Selling_Price', align: 'right', isNumber: true },
                          { key: 'UK_Size', align: 'center' },
                          { key: 'Euro_Size', align: 'center' },
                          { key: 'Color', align: 'left' },
                          { key: 'Product_Description', align: 'left', truncate: true },
                          { key: 'HSN_Code', align: 'center' },
                        ].map((cellMeta) => {
                          const colKey = cellMeta.key as keyof MasterRow;
                          const isEditing = editingCell?.id === prod.id && editingCell?.col === colKey;
                          const value = prod[colKey];

                          return (
                            <td 
                              key={colKey}
                              onClick={() => startEditing(prod.id, colKey, value)}
                              className={`py-2 px-3 border-r border-slate-100 cursor-pointer relative group/cell hover:bg-yellow-50/40 transition-colors ${
                                cellMeta.align === 'right' ? 'text-right' : cellMeta.align === 'center' ? 'text-center font-mono' : ''
                              } ${cellMeta.bold ? 'font-semibold text-slate-900 font-mono' : ''}`}
                            >
                              {isEditing ? (
                                <input
                                  type={cellMeta.isNumber ? "number" : "text"}
                                  value={editValue}
                                  onChange={(e) => setEditValue(e.target.value)}
                                  onBlur={saveCellEdit}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter') saveCellEdit();
                                    if (e.key === 'Escape') setEditingCell(null);
                                  }}
                                  autoFocus
                                  className="w-full bg-white border border-blue-500 rounded px-1.5 py-0.5 text-slate-800 font-medium focus:outline-hidden text-xs"
                                />
                              ) : (
                                <div className="flex items-center justify-between gap-1 w-full">
                                  <span className={cellMeta.truncate ? 'max-w-[160px] truncate block' : ''}>
                                    {cellMeta.isNumber && typeof value === 'number' ? `₹${value.toLocaleString()}` : String(value || '')}
                                  </span>
                                  <span className="text-[10px] text-slate-300 opacity-0 group-hover/cell:opacity-100 transition-opacity whitespace-nowrap">✏️</span>
                                </div>
                              )}
                            </td>
                          );
                        })}

                        {/* Image URLs indicators */}
                        <td className="py-2 px-3">
                          <div className="flex items-center gap-1">
                            <span 
                              className={`w-2 h-2 rounded-full ${prod.Main_Image_URL ? 'bg-emerald-500' : 'bg-slate-300'}`} 
                              title={prod.Main_Image_URL ? `Main URL: ${prod.Main_Image_URL}` : "Main URL missing"}
                            />
                            <span 
                              className={`w-2 h-2 rounded-full ${prod.Image_URL_2 ? 'bg-indigo-400' : 'bg-slate-200'}`} 
                              title={prod.Image_URL_2 ? `URL 2: ${prod.Image_URL_2}` : "URL 2 missing"}
                            />
                            <span 
                              className={`w-2 h-2 rounded-full ${prod.Image_URL_3 ? 'bg-amber-400' : 'bg-slate-200'}`} 
                              title={prod.Image_URL_3 ? `URL 3: ${prod.Image_URL_3}` : "URL 3 missing"}
                            />
                          </div>
                        </td>
                      </tr>
                    ))}

                    {filteredProducts.length === 0 && (
                      <tr>
                        <td colSpan={14} className="py-12 text-center text-slate-400 hover:bg-transparent">
                          <AlertCircle className="w-8 h-8 mx-auto text-slate-300 mb-2" />
                          <p className="text-sm font-semibold">No products matches query criteria</p>
                          <p className="text-xs text-slate-400 mt-1">Try resetting back to standard sample dataset</p>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* PAGINATION / FOOTER INFO */}
              <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-4">
                <span className="text-xs text-slate-500">
                  Showing {filteredProducts.length > 0 ? (currentPage - 1) * rowsPerPage + 1 : 0} to {Math.min(currentPage * rowsPerPage, filteredProducts.length)} of {filteredProducts.length} entries (Total Master Pool: {products.length})
                </span>

                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                    disabled={currentPage === 1}
                    className="px-3 py-1 bg-white border border-slate-300 rounded-lg text-xs font-semibold text-slate-600 disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-50 active:bg-slate-100 transition-colors"
                  >
                    Previous
                  </button>
                  {[...Array(totalPages)].map((_, i) => (
                    <button
                      key={i}
                      onClick={() => setCurrentPage(i + 1)}
                      className={`w-7 h-7 rounded-lg text-xs font-bold transition-all ${
                        currentPage === i + 1 
                          ? 'bg-blue-600 text-white shadow-xs' 
                          : 'bg-white border border-slate-300 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      {i + 1}
                    </button>
                  ))}
                  <button
                    onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                    disabled={currentPage === totalPages}
                    className="px-3 py-1 bg-white border border-slate-300 rounded-lg text-xs font-semibold text-slate-600 disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-50 active:bg-slate-100 transition-colors"
                  >
                    Next
                  </button>
                </div>
              </div>
            </div>

            {/* ACTION BULK CONVERTER TRIGGER */}
            <div className="bg-slate-900 border border-slate-800 text-white rounded-xl p-8 shadow-md relative overflow-hidden">
              <div className="absolute right-0 top-0 translate-x-12 -translate-y-12 w-64 h-64 bg-slate-800/50 rounded-full blur-3xl pointer-events-none" />
              <div className="relative flex flex-col md:flex-row md:items-center md:justify-between gap-6">
                <div>
                  <h3 className="text-base font-bold text-white flex items-center gap-2">
                    🛠️ Transform Master Dataset
                  </h3>
                  <p className="text-xs text-slate-300 mt-1 max-w-2xl leading-relaxed">
                    Instantly transform the active workbench table state into three separate spreadsheets. Custom business rules are injected, such as force constant UK shoe sizing standard labels on Amazon and ISP mapped margins on Myntra.
                  </p>
                </div>

                <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3 shrink-0">
                  <button
                    onClick={handleBulkExportAll}
                    disabled={products.length === 0}
                    className="bg-blue-600 hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold text-xs py-3 px-6 rounded-xl shadow-lg shadow-blue-600/10 active:scale-[0.98] transition-all flex items-center justify-center gap-2"
                  >
                    <Download className="w-4.5 h-4.5" />
                    Download All Marketplace Templates (3)
                  </button>
                </div>
              </div>

              {/* THREE GRID PREVIEW TEMPLATES OR ACTION DOWNLOADS */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mt-8 pt-8 border-t border-slate-800">
                
                {/* AMAZON CARD */}
                <div className="bg-slate-805/50 border border-slate-800 rounded-xl p-5 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <span className="bg-amber-500/15 text-amber-300 font-bold px-2 py-0.5 rounded text-[10px] tracking-wider uppercase border border-amber-500/10">Amazon 🅰️</span>
                      <span className="text-[11px] text-slate-400">12 Mapped Headers</span>
                    </div>
                    <ul className="text-[11px] text-slate-400 space-y-2 mb-6">
                      <li className="flex items-start gap-1">
                        <span className="text-emerald-500 mr-1">✓</span>
                        <span>Size standard automatically set to <code className="text-white text-[10px] bg-slate-800 px-1 py-0.2 rounded font-mono">"UK"</code></span>
                      </li>
                      <li className="flex items-start gap-1">
                        <span className="text-emerald-500 mr-1">✓</span>
                        <span>Model attributes fall back on Style Code / Article Code</span>
                      </li>
                    </ul>
                  </div>

                  <button
                    onClick={downloadAmazon}
                    disabled={products.length === 0}
                    className="w-full bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-white font-bold py-2 px-4 rounded-lg text-xs transition-colors flex items-center justify-center gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Amazon_Bulk_Upload.xlsx
                  </button>
                </div>

                {/* FLIPKART CARD */}
                <div className="bg-slate-805/50 border border-slate-800 rounded-xl p-5 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <span className="bg-sky-500/15 text-sky-300 font-bold px-2 py-0.5 rounded text-[10px] tracking-wider uppercase border border-sky-500/10">Flipkart 🅵</span>
                      <span className="text-[11px] text-slate-400">14 Mapped Headers</span>
                    </div>
                    <ul className="text-[11px] text-slate-400 space-y-2 mb-6">
                      <li className="flex items-start gap-1">
                        <span className="text-emerald-500 mr-1">✓</span>
                        <span>Separated UK Size and Euro Size column exports</span>
                      </li>
                      <li className="flex items-start gap-1">
                        <span className="text-emerald-500 mr-1">✓</span>
                        <span>Mapex multi-angle image properties of listings</span>
                      </li>
                    </ul>
                  </div>

                  <button
                    onClick={downloadFlipkart}
                    disabled={products.length === 0}
                    className="w-full bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-white font-bold py-2 px-4 rounded-lg text-xs transition-colors flex items-center justify-center gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Flipkart_Bulk_Upload.xlsx
                  </button>
                </div>

                {/* MYNTRA CARD */}
                <div className="bg-slate-805/50 border border-slate-800 rounded-xl p-5 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <span className="bg-pink-500/15 text-pink-300 font-bold px-2 py-0.5 rounded text-[10px] tracking-wider uppercase border border-pink-500/10">Myntra 🅼</span>
                      <span className="text-[11px] text-slate-400">14 Mapped Headers</span>
                    </div>
                    <ul className="text-[11px] text-slate-400 space-y-2 mb-6">
                      <li className="flex items-start gap-1">
                        <span className="text-emerald-500 mr-1">✓</span>
                        <span>Selling_Price maps to <code className="text-white text-[10px] bg-slate-800 px-1 py-0.2 rounded font-mono">"ISP"</code> header</span>
                      </li>
                      <li className="flex items-start gap-1">
                        <span className="text-emerald-500 mr-1">✓</span>
                        <span>Front, Side, and Back images parsed cleanly</span>
                      </li>
                    </ul>
                  </div>

                  <button
                    onClick={downloadMyntra}
                    disabled={products.length === 0}
                    className="w-full bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-white font-bold py-2 px-4 rounded-lg text-xs transition-colors flex items-center justify-center gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Myntra_Bulk_Upload.xlsx
                  </button>
                </div>

              </div>
            </div>

          </div>
        )}

        {/* STREAMLIT PORTABILITY PYTHON TAB */}
        {activeTab === 'portability' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs">
              <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 mb-4">
                <div>
                  <h2 className="text-base font-bold text-slate-800">🔌 Streamlit + Pandas Portable Application (app.py)</h2>
                  <p className="text-xs text-slate-500 mt-0.5">Use this exact script on your computer desktop with Python virtual environments anytime.</p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={copyToClipboard}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-xs py-1.5 px-3 rounded-lg border border-slate-300 transition-colors flex items-center gap-1"
                  >
                    {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
                    {copied ? 'Copied' : 'Copy Python Code'}
                  </button>

                  <button
                    onClick={downloadPythonScript}
                    className="bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs py-1.5 px-3.5 rounded-lg transition-colors flex items-center gap-1 shadow-sm"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Download app.py Script
                  </button>
                </div>
              </div>

              {/* HANDOVER PACKAGE SECTION */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-6 mb-6 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-4 pb-4 border-b border-slate-200">
                  <div>
                    <h3 className="font-bold text-slate-900 text-sm flex items-center gap-1.5">
                      <span className="p-1 bg-blue-100 text-blue-700 rounded-lg text-xs font-bold">📩</span>
                      Direct Offline Handover Package for Employees
                    </h3>
                    <p className="text-[11px] text-slate-500 mt-1">
                      Download these files, attach them to an email, and send them directly to your employee!
                    </p>
                  </div>
                  <div className="text-[10px] text-slate-500 bg-slate-100 px-2.5 py-1 rounded-md border border-slate-200 font-medium self-start">
                    ⚖️ Desktop Native Solution
                  </div>
                </div>

                <div className="text-[11px] text-slate-650 leading-relaxed space-y-3">
                  <p>
                    Because this web application processes data in a secure, isolated cloud environment running on a Linux pipeline, it cannot cross-compile customized native Windows <code className="bg-slate-100 px-1 py-0.5 rounded font-mono decoration-slate-300">.exe</code> executables on-the-fly due to operating system differences and browser-security restrictions.
                  </p>
                  <p className="font-medium text-slate-800">
                    💡 However, we have solved this for you perfectly! By sending your employee this simple 3-file package, they can double-click to launch the app instantly offline on their Windows computer without ever seeing any terminal commands or code:
                  </p>
                </div>

                {/* 3 FILES BENTO GRID */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-1">
                  
                  {/* FILE 1: APPLICATION SCRIPT */}
                  <div className="bg-white border border-slate-200 hover:border-blue-200 rounded-lg p-4 flex flex-col justify-between transition-all shadow-xs">
                    <div>
                      <div className="flex items-center gap-1.5 mb-2">
                        <span className="text-blue-500 text-xs">📄</span>
                        <h4 className="font-bold text-slate-800 text-[11px]">1. Main Python Script</h4>
                      </div>
                      <p className="text-[10px] text-slate-500 leading-relaxed mb-4">
                        The core footwear list processor incorporating all Amazon, Flipkart, & Myntra schemas.
                      </p>
                    </div>
                    <div>
                      <div className="font-mono text-[9px] text-slate-400 mb-2 bg-slate-50 px-1.5 py-0.5 rounded inline-block">
                        File: app.py
                      </div>
                      <button
                        onClick={downloadPythonScript}
                        className="w-full bg-blue-50 hover:bg-blue-100 text-blue-700 font-bold py-2 px-3 rounded-md text-xs transition-colors flex items-center justify-center gap-1.5 border border-blue-200 shadow-2xs"
                      >
                        <Download className="w-3 h-3" />
                        Download app.py
                      </button>
                    </div>
                  </div>

                  {/* FILE 2: WINDOWS LAUNCHER */}
                  <div className="bg-white border border-slate-200 hover:border-emerald-200 rounded-lg p-4 flex flex-col justify-between transition-all shadow-xs">
                    <div>
                      <div className="flex items-center gap-1.5 mb-2">
                        <span className="text-emerald-500 text-xs">⚙️</span>
                        <h4 className="font-bold text-slate-800 text-[11px]">2. Windows Auto-Launcher</h4>
                      </div>
                      <p className="text-[10px] text-slate-500 leading-relaxed mb-4">
                        A custom executable script. Just double-click it to verify Python, install modules, and open the app!
                      </p>
                    </div>
                    <div>
                      <div className="font-mono text-[9px] text-emerald-600 mb-2 bg-emerald-50 px-1.5 py-0.5 rounded inline-block">
                        File: Launch_Footwear_App.bat
                      </div>
                      <button
                        onClick={downloadBatLauncher}
                        className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 px-3 rounded-md text-xs transition-colors flex items-center justify-center gap-1.5 shadow-sm"
                      >
                        <Download className="w-3 h-3" />
                        Download Launcher
                      </button>
                    </div>
                  </div>

                  {/* FILE 3: EMPLOYEE MANUAL */}
                  <div className="bg-white border border-slate-200 hover:border-amber-200 rounded-lg p-4 flex flex-col justify-between transition-all shadow-xs">
                    <div>
                      <div className="flex items-center gap-1.5 mb-2">
                        <span className="text-amber-500 text-xs">📖</span>
                        <h4 className="font-bold text-slate-800 text-[11px]">3. Handover Manual</h4>
                      </div>
                      <p className="text-[10px] text-slate-500 leading-relaxed mb-4">
                        A simple, step-by-step instructions text guide written directly for your employee.
                      </p>
                    </div>
                    <div>
                      <div className="font-mono text-[9px] text-amber-600 mb-2 bg-amber-50 px-1.5 py-0.5 rounded inline-block">
                        File: Handover_Manual.txt
                      </div>
                      <button
                        onClick={downloadReadmeInstructions}
                        className="w-full bg-amber-50 hover:bg-amber-100 text-amber-805 font-bold py-2 px-3 rounded-md text-xs transition-colors flex items-center justify-center gap-1.5 border border-amber-200 shadow-2xs"
                      >
                        <Download className="w-3 h-3" />
                        Download Manual
                      </button>
                    </div>
                  </div>

                </div>

                <div className="p-3.5 bg-blue-50/50 border border-blue-100 rounded-lg text-[11px] text-blue-900 leading-relaxed">
                  📧 <strong>How to handover:</strong> Simply download all three files, attach them to an email to your employee, and copy-paste the contents of the <em>Handover Manual</em> into the body of the email.
                </div>
              </div>

              {/* LOCAL OFFLINE DESKTOP EXE PACKAGING CARD */}
              <div className="bg-emerald-50/50 border border-emerald-200 rounded-xl p-5 text-xs text-slate-800 mb-6 space-y-4">
                <div className="flex items-center gap-2 pb-2 border-b border-emerald-100">
                  <div className="p-1 bg-emerald-600 rounded-lg text-white font-bold text-sm">
                    📦
                  </div>
                  <div>
                    <h3 className="font-bold text-slate-900 text-sm">How to compile a Standalone .exe (Optional)</h3>
                    <p className="text-[11px] text-slate-500">Run this compilation tool on a Windows computer to bundle everything into a single app file</p>
                  </div>
                </div>

                <p className="leading-relaxed">
                  If your employee prefers a native Windows compiled <code className="bg-emerald-100/80 font-mono text-emerald-850 px-1 py-0.2 rounded text-[11px]">.exe</code> binary file directly instead of running the launcher, they can easily build one themselves locally inside their command terminal by running standard commands:
                </p>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 font-sans">
                  <div>
                    <span className="font-semibold text-slate-800 text-xs block mb-1">A. Install the Packaging Tool</span>
                    <pre className="bg-slate-900 text-slate-100 p-3 rounded-lg text-[11px] font-mono select-all overflow-x-auto">
                      pip install pyinstaller
                    </pre>
                  </div>
                  <div>
                    <span className="font-semibold text-slate-800 text-xs block mb-1">B. Convert Script to Standalone App</span>
                    <pre className="bg-slate-900 text-slate-100 p-3 rounded-lg text-[11px] font-mono select-all overflow-x-auto">
                      pyinstaller --onefile --hidden-import=streamlit app.py
                    </pre>
                  </div>
                </div>

                <div className="text-[11px] text-slate-650 leading-relaxed pt-1">
                  Once compiled, open the newly created <code className="bg-slate-100 px-1 py-0.5 rounded font-mono text-slate-700">dist</code> folder on that computer. Inside, they will find an <strong className="text-slate-800 font-semibold">app.exe</strong> file which can be shared, double-clicked, and run completely offline!
                </div>
              </div>

              {/* TEXT AREA CODE BLOCK DISPLAY */}
              <div className="relative">
                <div className="absolute top-3 right-3 bg-slate-800 text-slate-400 font-mono text-[9px] px-2 py-0.5 rounded">
                  PYTHON
                </div>
                <textarea
                  readOnly
                  value={streamLitCodeStr}
                  className="w-full h-96 bg-slate-950 text-slate-100 font-mono text-xs rounded-xl p-5 focus:outline-hidden border border-slate-800 shadow-inner resize-y select-text leading-relaxed"
                />
              </div>
            </div>
          </div>
        )}

        {/* REFERENCE INFO TAB */}
        {activeTab === 'reference' && (
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            
            {/* AMAZON INFO */}
            <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col justify-between">
              <div>
                <span className="bg-amber-100 text-amber-800 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase tracking-wider mb-3 inline-block">Amazon 🅰️</span>
                <h3 className="text-sm font-bold text-slate-800 mb-2">Target Schema Translation</h3>
                <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                  Amazon bulk templates enforce specific footwear system constraints which this utility automatically executes.
                </p>

                <div className="space-y-2 border-t border-slate-100 pt-4 text-xs">
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">Footwear Size System</span>
                    <span className="font-semibold text-slate-800 bg-slate-100 px-1.5 py-0.2 rounded">"UK" (Constant)</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">Shoe Size</span>
                    <span className="font-semibold text-slate-800 font-mono">UK_Size</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">Model number</span>
                    <span className="text-slate-500 text-right">Style_Code / Article_Number</span>
                  </div>
                  <div className="flex justify-between py-1">
                    <span className="text-slate-400 font-mono">Your Price</span>
                    <span className="font-semibold text-slate-800 font-mono">Selling_Price</span>
                  </div>
                </div>
              </div>
            </div>

            {/* FLIPKART INFO */}
            <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col justify-between">
              <div>
                <span className="bg-sky-100 text-sky-800 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase tracking-wider mb-3 inline-block">Flipkart 🅵</span>
                <h3 className="text-sm font-bold text-slate-800 mb-2">Target Schema Translation</h3>
                <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                  Flipkart templates accept independent sizes (UK and Euro structures) alongside nested other asset URLs.
                </p>

                <div className="space-y-2 border-t border-slate-100 pt-4 text-xs">
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">Seller SKU ID</span>
                    <span className="font-semibold text-slate-800 font-mono">Master_SKU</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">UK/India Size</span>
                    <span className="font-semibold text-slate-800 font-mono">UK_Size</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">Euro Size</span>
                    <span className="font-semibold text-slate-800 font-mono">Euro_Size</span>
                  </div>
                  <div className="flex justify-between py-1">
                    <span className="text-slate-400 font-mono">Other Image URL 1/2</span>
                    <span className="text-slate-500 text-right">Image_URL_2 / Image_URL_3</span>
                  </div>
                </div>
              </div>
            </div>

            {/* MYNTRA INFO */}
            <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col justify-between">
              <div>
                <span className="bg-pink-100 text-pink-800 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase tracking-wider mb-3 inline-block">Myntra 🅼</span>
                <h3 className="text-sm font-bold text-slate-800 mb-2">Target Schema Translation</h3>
                <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                  Myntra relies on unique brand keys ("ISP" for Selling Price, "vendorSkuCode" for SKU, and Front/Side/Back image specs).
                </p>

                <div className="space-y-2 border-t border-slate-100 pt-4 text-xs">
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">vendorSkuCode</span>
                    <span className="font-semibold text-slate-800 font-mono">Master_SKU</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">ISP</span>
                    <span className="font-semibold text-slate-850 font-mono bg-pink-50 px-1 py-0.2 text-pink-800 rounded">Selling_Price</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-slate-50">
                    <span className="text-slate-400 font-mono">Prominent Colour</span>
                    <span className="font-semibold text-slate-800">Color</span>
                  </div>
                  <div className="flex justify-between py-1">
                    <span className="text-slate-400 font-mono">Product Details</span>
                    <span className="text-slate-500">Product_Description</span>
                  </div>
                </div>
              </div>
            </div>

          </div>
        )}

      </main>

      {/* FOOTER */}
      <footer className="border-t border-slate-200 bg-white py-6 mt-12 text-center text-xs text-slate-400 font-mono uppercase tracking-wider">
        <span>© 2026 Footwear Bulk Listing Generator • Powered by React & SheetJS</span>
      </footer>

    </div>
  );
}
