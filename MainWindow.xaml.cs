//MainWindow.xaml.cs
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Globalization;
using System.Threading;

namespace CocktailViewer
{
    public abstract class LibraryNode
    {
        public required string Name { get; init; }
    }

    public sealed class FolderNode : LibraryNode
    {
        public ObservableCollection<LibraryNode> Children { get; } = new();
        public string? FullPath { get; init; }
    }

    public sealed class ImageNode : LibraryNode, INotifyPropertyChanged
    {
        public required string FullPath { get; init; }

        private BitmapSource? _thumbnail;
        public BitmapSource? Thumbnail
        {
            get => _thumbnail;
            set
            {
                if (ReferenceEquals(_thumbnail, value)) return;
                _thumbnail = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        // Open images (right side)
        public ObservableCollection<ImageItem> Images { get; } = new();

        // Folder tree (left side)
        public ObservableCollection<FolderNode> LibraryRoot { get; } = new();

        private readonly Dictionary<string, ImageItem> _openByPath =
            new(StringComparer.OrdinalIgnoreCase);

        // Track which folders have had thumbnails loaded (lazy loading)
        private readonly HashSet<string> _thumbsLoadedForFolder =
            new(StringComparer.OrdinalIgnoreCase);
			
		private readonly HashSet<string> _thumbnailsLoading =
			new(StringComparer.OrdinalIgnoreCase);
			
		private readonly HashSet<string> _expandedFolders =
			new(StringComparer.OrdinalIgnoreCase);

        // Touch debounce lock (folder header taps)
        private bool _folderClickLocked;
        private readonly DispatcherTimer _folderClickTimer;
        private static readonly TimeSpan FolderClickDebounce = TimeSpan.FromMilliseconds(180);

        // Open-image drag/reorder state
        private Point _openImageDragStartPoint;
        private ImageItem? _draggedOpenImage;
        private bool _openImageDragPending;
        private bool _suppressNextImageToggleClick;
		private ImageItem? _lastDragOverTarget;

        // Supported extensions
        private static readonly HashSet<string> SupportedExt =
            new(StringComparer.OrdinalIgnoreCase)
            { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".jfif" };
			
        public sealed class ImageItem : INotifyPropertyChanged
        {
            public required string Path { get; init; }
            public required BitmapSource Bitmap { get; init; }

            private bool _isBeingDragged;
            public bool IsBeingDragged
            {
                get => _isBeingDragged;
                set
                {
                    if (_isBeingDragged == value)
                        return;

                    _isBeingDragged = value;
                    OnPropertyChanged();
                }
            }

            public event PropertyChangedEventHandler? PropertyChanged;
            private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // Design-space height so all open images scale identically
        private double _designHeight = 1080;
        public double DesignHeight
        {
            get => _designHeight;
            set
            {
                if (Math.Abs(_designHeight - value) < 0.5) return;
                _designHeight = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // Saved library-folder config
        private static readonly string ConfigFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CocktailViewer");

        private static readonly string LibraryPathFile =
            Path.Combine(ConfigFolder, "librarypath.txt");

        private string _libraryFolder = GetInitialLibraryFolder();

        private string LibraryFolder => _libraryFolder;

		private static string GetInitialLibraryFolder()
		{
			try
			{
				if (File.Exists(LibraryPathFile))
				{
					var saved = File.ReadAllText(LibraryPathFile).Trim();

					if (!string.IsNullOrWhiteSpace(saved))
					{
						if (Directory.Exists(saved))
							return saved;

						MessageBox.Show(
							"Unable to find the saved library folder.\n\nReverting to the default library location.",
							"CocktailViewer",
							MessageBoxButton.OK,
							MessageBoxImage.Warning);
					}
				}
			}
			catch
			{
				// Ignore errors and fall back
			}

			return Path.Combine(Path.GetDirectoryName(Environment.ProcessPath!)!, "Library");
		}

        // Live watching
        private FileSystemWatcher? _libraryWatcher;
        private DispatcherTimer? _libraryDebounceTimer;
        private static readonly TimeSpan LibraryDebounce = TimeSpan.FromMilliseconds(250);
		
		// Incremental refresh tracking
		private readonly HashSet<string> _dirtyFolders = new(StringComparer.OrdinalIgnoreCase);
		
		private static readonly string ThumbnailCacheFolder =
			Path.Combine(ConfigFolder, "thumbcache");

		private readonly SemaphoreSlim _thumbnailDecodeGate = new(4); // max 4 concurrent decodes
		
		// Periodic polling backup
		private DispatcherTimer? _libraryPollTimer;
		private static readonly TimeSpan LibraryPollInterval = TimeSpan.FromSeconds(10);

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            WindowState = WindowState.Maximized;

            _folderClickTimer = new DispatcherTimer { Interval = FolderClickDebounce };
            _folderClickTimer.Tick += (_, __) =>
            {
                _folderClickLocked = false;
                _folderClickTimer.Stop();
            };

            Loaded += (_, __) =>
            {
                EnsureLibraryFolder();
                FitToScreen();
                UpdateDesignHeight();

                BuildLibraryTreeSkeleton();        // fast: builds folder + image nodes (no subfolder thumbnails)
                LoadThumbnailsForRootOnly();       // load thumbnails only for images directly under Library\
                EnsureRootExpanded();              // root always expanded
                StartLibraryWatcher();
				//StartLibraryPolling();

                // Command-line images (copy into library, then open)
                var args = Environment.GetCommandLineArgs();
                if (args.Length >= 2)
                {
                    foreach (var p in args.Skip(1).Where(File.Exists))
                        OpenImageFromAnywhere(p);
                }

				Directory.CreateDirectory(ThumbnailCacheFolder);
					_ = Task.Run(SweepThumbnailCache);

                UpdateHint();

            };

            Closed += (_, __) =>
            {
                StopLibraryWatcher();
				//StopLibraryPolling();
                _folderClickTimer.Stop();
            };
        }

		private static string ComputeSha256Hex(string text)
		{
			using var sha = SHA256.Create();
			var bytes = Encoding.UTF8.GetBytes(text);
			var hash = sha.ComputeHash(bytes);
			return Convert.ToHexString(hash).ToLowerInvariant();
		}

		private static string GetThumbnailCacheKey(string imagePath)
		{
			try
			{
				var fi = new FileInfo(imagePath);
				string fullPath = Path.GetFullPath(imagePath);
				string payload = string.Create(CultureInfo.InvariantCulture,
					$"{fullPath}|{fi.LastWriteTimeUtc.Ticks}|{fi.Length}");
				return ComputeSha256Hex(payload);
			}
			catch
			{
				return ComputeSha256Hex(Path.GetFullPath(imagePath));
			}
		}

		private static string GetThumbnailCachePath(string imagePath)
		{
			var key = GetThumbnailCacheKey(imagePath);
			return Path.Combine(ThumbnailCacheFolder, key + ".png");
		}

		private static BitmapSource? LoadBitmapSourceFromFile(string path)
		{
			try
			{
				if (!File.Exists(path))
					return null;

				var bi = new BitmapImage();
				bi.BeginInit();
				bi.CacheOption = BitmapCacheOption.OnLoad;
				bi.UriSource = new Uri(path, UriKind.Absolute);
				bi.EndInit();
				bi.Freeze();
				return bi;
			}
			catch
			{
				return null;
			}
		}

		private static BitmapSource? CreateThumbnailAndPersistCache(string sourcePath, string cachePath)
		{
			try
			{
				var thumb = LoadThumbnail(sourcePath);
				if (thumb == null)
					return null;

				Directory.CreateDirectory(ThumbnailCacheFolder);

				var encoder = new PngBitmapEncoder();
				encoder.Frames.Add(BitmapFrame.Create(thumb));

				using (var fs = new FileStream(cachePath, FileMode.Create, FileAccess.Write, FileShare.None))
				{
					encoder.Save(fs);
				}

				return thumb;
			}
			catch
			{
				return null;
			}
		}
		
		private void SweepThumbnailCache()
		{
			try
			{
				Directory.CreateDirectory(ThumbnailCacheFolder);

				var valid = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

				IEnumerable<string> files;
				try
				{
					files = Directory.EnumerateFiles(LibraryFolder, "*.*", SearchOption.AllDirectories)
						.Where(f => SupportedExt.Contains(Path.GetExtension(f)));
				}
				catch
				{
					return;
				}

				foreach (var file in files)
				{
					valid.Add(GetThumbnailCachePath(file));
				}

				foreach (var cachedFile in Directory.EnumerateFiles(ThumbnailCacheFolder, "*.png", SearchOption.TopDirectoryOnly))
				{
					if (!valid.Contains(cachedFile))
					{
						try { File.Delete(cachedFile); } catch { }
					}
				}
			}
			catch
			{
				// ignore
			}
		}

        private void EnsureLibraryFolder()
        {
            Directory.CreateDirectory(LibraryFolder);
        }

		private static void SaveLibraryFolder(string folderPath)
		{
			try
			{
				Directory.CreateDirectory(ConfigFolder);
				File.WriteAllText(LibraryPathFile, folderPath);
			}
			catch
			{
				// intentionally ignore save errors
			}
		}

        private void FitToScreen()
        {
            var wa = SystemParameters.WorkArea;
            Left = wa.Left;
            Top = wa.Top;
            Width = wa.Width;
            Height = wa.Height;
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateDesignHeight();
        }

        private void UpdateDesignHeight()
        {
            DesignHeight = Math.Max(1, ActualHeight - 2);
        }

        private void UpdateHint()
        {
            HintText.Visibility = Images.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        // Root expands after containers exist
        private void EnsureRootExpanded()
        {
            if (LibraryRoot.Count == 0) return;

            var root = LibraryRoot[0];

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (LibraryTree.ItemContainerGenerator.ContainerFromItem(root) is TreeViewItem tvi)
                {
                    tvi.IsExpanded = true;
                }
            }), DispatcherPriority.Loaded);
        }

        // ✅ Browse Library button handler
        private void BrowseLibrary_Click(object sender, RoutedEventArgs e)
        {
            EnsureLibraryFolder();

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{LibraryFolder}\"",
                    UseShellExecute = true
                });
            }
            catch
            {
                // intentionally ignore (no UI change requested)
            }
        }

        private void SetLibraryLocation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new OpenFolderDialog
                {
                    Title = "Select Library Folder",
                    InitialDirectory = Directory.Exists(LibraryFolder)
                        ? LibraryFolder
                        : Path.GetDirectoryName(Environment.ProcessPath!)!
                };

                if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.FolderName))
                {
                    SetLibraryFolder(dialog.FolderName);
                }
            }
            catch
            {
                // Ignore for now
            }
        }

		private void SetLibraryFolder(string folderPath)
		{
			if (string.IsNullOrWhiteSpace(folderPath))
				return;

			folderPath = Path.GetFullPath(folderPath);

			Directory.CreateDirectory(folderPath);

			StopLibraryWatcher();
			//StopLibraryPolling();

			_libraryFolder = folderPath;
			SaveLibraryFolder(folderPath);

			BuildLibraryTreeSkeleton();
			LoadThumbnailsForRootOnly();
			EnsureRootExpanded();
			StartLibraryWatcher();
			//StartLibraryPolling();
			UpdateHint();
		}

		private void MinimizeWindow_Click(object sender, RoutedEventArgs e)
		{
			WindowState = WindowState.Minimized;
		}

		private void AlignImagesLeft_Click(object sender, RoutedEventArgs e)
		{
			CocktailViewbox.HorizontalAlignment = HorizontalAlignment.Left;
		}

		private void AlignImagesCenter_Click(object sender, RoutedEventArgs e)
		{
			CocktailViewbox.HorizontalAlignment = HorizontalAlignment.Center;
		}
		
	    // ---- Drag & drop ----

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0)
                return;

            foreach (var f in files.Where(File.Exists))
                OpenImageFromAnywhere(f);
        }

        // ---- Open from anywhere (drop/cmdline) ----

        private void OpenImageFromAnywhere(string sourcePath)
        {
            if (!File.Exists(sourcePath))
                return;

            var ext = Path.GetExtension(sourcePath);
            if (!SupportedExt.Contains(ext))
                return;

            var libraryPath = CopyIntoLibrary(sourcePath);

            // Rebuild skeleton & root thumbs (simple and consistent)
            BuildLibraryTreeSkeleton();
            LoadThumbnailsForRootOnly();
            EnsureRootExpanded();

            OpenLibraryImage(libraryPath);
        }

        // ---- Copying & naming ----

        private string CopyIntoLibrary(string sourcePath)
        {
            var ext = Path.GetExtension(sourcePath);
            var baseName = Path.GetFileNameWithoutExtension(sourcePath);

            string hash8 = ComputeFileHash8(sourcePath);
            string safeBase = MakeSafeFileName(baseName);

            var destName = $"{safeBase}_{hash8}{ext}";
            var destPath = Path.Combine(LibraryFolder, destName);

            if (!File.Exists(destPath))
                File.Copy(sourcePath, destPath);

            return destPath;
        }

        private static string ComputeFileHash8(string path)
        {
            using var stream = File.OpenRead(path);
            using var sha = SHA256.Create();
            var hash = sha.ComputeHash(stream);
            return BitConverter.ToString(hash, 0, 4).Replace("-", "").ToLowerInvariant();
        }

        private static string MakeSafeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name)
                sb.Append(invalid.Contains(ch) ? '_' : ch);

            var s = sb.ToString().Trim();
            return string.IsNullOrWhiteSpace(s) ? "image" : s;
        }

        // ---- Live watching ----

		private void StartLibraryWatcher()
		{
			StopLibraryWatcher();

			_libraryDebounceTimer = new DispatcherTimer { Interval = LibraryDebounce };
			_libraryDebounceTimer.Tick += (_, __) =>
			{
				_libraryDebounceTimer!.Stop();
				RefreshDirtyFolders();
			};

			_libraryWatcher = new FileSystemWatcher(LibraryFolder)
			{
				IncludeSubdirectories = true,
				Filter = "*.*",
				NotifyFilter = NotifyFilters.FileName
							 | NotifyFilters.DirectoryName
							 | NotifyFilters.LastWrite
							 | NotifyFilters.Size
							 | NotifyFilters.CreationTime,
				InternalBufferSize = 64 * 1024
			};

			_libraryWatcher.Created += (_, e) => ScheduleLibraryRefreshForPath(e.FullPath);
			_libraryWatcher.Deleted += (_, e) => ScheduleLibraryRefreshForPath(e.FullPath);
			_libraryWatcher.Changed += (_, e) => ScheduleLibraryRefreshForPath(e.FullPath);
			_libraryWatcher.Renamed += (_, e) =>
			{
				ScheduleLibraryRefreshForPath(e.OldFullPath);
				ScheduleLibraryRefreshForPath(e.FullPath);
			};

			_libraryWatcher.Error += (_, __) =>
			{
				Dispatcher.Invoke(() =>
				{
					RefreshLibraryFromDisk();
					StopLibraryWatcher();
					StartLibraryWatcher();
				});
			};

			_libraryWatcher.EnableRaisingEvents = true;
		}

        private void StopLibraryWatcher()
        {
            if (_libraryWatcher != null)
            {
                _libraryWatcher.EnableRaisingEvents = false;
                _libraryWatcher.Dispose();
                _libraryWatcher = null;
            }

            if (_libraryDebounceTimer != null)
            {
                _libraryDebounceTimer.Stop();
                _libraryDebounceTimer = null;
            }
        }
		
		private void RefreshDirtyFolders()
		{
			if (LibraryRoot.Count == 0)
				return;

			if (_dirtyFolders.Count == 0)
			{
				RefreshLibraryFromDisk();
				return;
			}

			var dirty = _dirtyFolders.ToList();
			_dirtyFolders.Clear();

			foreach (var folderPath in dirty)
				RefreshSingleFolder(folderPath);

			// Keep root expanded and root thumbnails available.
			EnsureRootExpanded();
			LoadThumbnailsForRootOnly();
		}
		
		private void RefreshSingleFolder(string folderPath)
		{
			if (LibraryRoot.Count == 0)
				return;

			folderPath = Path.GetFullPath(folderPath);

			// If the change is outside the current library somehow, just do a full refresh.
			if (!IsUnderLibrary(folderPath))
			{
				RefreshLibraryFromDisk();
				return;
			}

			var root = LibraryRoot[0];

			// Refreshing the library root is effectively a full rebuild.
			if (string.Equals(folderPath, LibraryFolder, StringComparison.OrdinalIgnoreCase))
			{
				RefreshLibraryFromDisk();
				return;
			}

			var parentPath = Path.GetDirectoryName(folderPath);
			if (string.IsNullOrWhiteSpace(parentPath))
				parentPath = LibraryFolder;

			var parentNode = FindFolderNode(root, parentPath);
			if (parentNode == null)
			{
				RefreshLibraryFromDisk();
				return;
			}

			// Remove existing node for this folder from parent, if present.
			var existingFolder = parentNode.Children
				.OfType<FolderNode>()
				.FirstOrDefault(f => string.Equals(f.FullPath, folderPath, StringComparison.OrdinalIgnoreCase));

			if (existingFolder != null)
			{
				RemoveThumbStateRecursive(existingFolder);
				parentNode.Children.Remove(existingFolder);
			}

			// Re-add it if it still exists on disk.
			if (Directory.Exists(folderPath))
			{
				var rebuiltFolder = BuildFolderNodeFromDisk(folderPath);

				if (rebuiltFolder != null)
				{
					parentNode.Children.Add(rebuiltFolder);

					// If that folder had been expanded previously, eagerly reload its thumbnails.
					if (_thumbsLoadedForFolder.Contains(folderPath))
						LoadThumbnailsForFolderForce(rebuiltFolder);
				}
			}

			SortFolderRecursive(parentNode);

			// Also refresh the parent itself in case file adds/deletes happened directly in it.
			RefreshFolderImageChildrenFromDisk(parentNode);
			SortFolderRecursive(parentNode);
		}

		private bool IsUnderLibrary(string path)
		{
			var library = Path.GetFullPath(LibraryFolder)
				.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
				+ Path.DirectorySeparatorChar;

			var candidate = Path.GetFullPath(path)
				.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
				+ Path.DirectorySeparatorChar;

			return candidate.StartsWith(library, StringComparison.OrdinalIgnoreCase)
				|| string.Equals(
					Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
					Path.GetFullPath(LibraryFolder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
					StringComparison.OrdinalIgnoreCase);
		}

		private FolderNode? FindFolderNode(FolderNode current, string folderPath)
		{
			if (current.FullPath != null &&
				string.Equals(current.FullPath, folderPath, StringComparison.OrdinalIgnoreCase))
				return current;

			foreach (var child in current.Children.OfType<FolderNode>())
			{
				var found = FindFolderNode(child, folderPath);
				if (found != null)
					return found;
			}

			return null;
		}

		private FolderNode? BuildFolderNodeFromDisk(string folderPath)
		{
			if (!Directory.Exists(folderPath))
				return null;

			var node = new FolderNode
			{
				Name = Path.GetFileName(folderPath),
				FullPath = folderPath
			};

			try
			{
				foreach (var dir in Directory.EnumerateDirectories(folderPath))
				{
					var child = BuildFolderNodeFromDisk(dir);
					if (child != null)
						node.Children.Add(child);
				}

				foreach (var file in Directory.EnumerateFiles(folderPath))
				{
					if (!SupportedExt.Contains(Path.GetExtension(file)))
						continue;

					node.Children.Add(new ImageNode
					{
						Name = Path.GetFileName(file),
						FullPath = file,
						Thumbnail = null
					});
				}
			}
			catch
			{
				return node;
			}

			SortFolderRecursive(node);
			return node;
		}

		private void RefreshFolderImageChildrenFromDisk(FolderNode folder)
		{
			if (folder.FullPath == null || !Directory.Exists(folder.FullPath))
				return;

			var existingSubfolders = folder.Children.OfType<FolderNode>().ToList();
			folder.Children.Clear();

			foreach (var subfolder in existingSubfolders)
				folder.Children.Add(subfolder);

			try
			{
				foreach (var file in Directory.EnumerateFiles(folder.FullPath))
				{
					if (!SupportedExt.Contains(Path.GetExtension(file)))
						continue;

					folder.Children.Add(new ImageNode
					{
						Name = Path.GetFileName(file),
						FullPath = file,
						Thumbnail = null
					});
				}
			}
			catch
			{
				// ignore
			}

			if (_thumbsLoadedForFolder.Contains(folder.FullPath))
				LoadThumbnailsForFolderForce(folder);
		}

		private void LoadThumbnailsForFolderForce(FolderNode folder)
		{
			if (folder.FullPath == null)
				return;

			_thumbsLoadedForFolder.Add(folder.FullPath);

			foreach (var img in folder.Children.OfType<ImageNode>())
			{
				if (img.Thumbnail != null) continue;
				_ = LoadThumbnailAsync(img);
			}
		}

		private void RemoveThumbStateRecursive(FolderNode folder)
		{
			if (!string.IsNullOrWhiteSpace(folder.FullPath))
				_thumbsLoadedForFolder.Remove(folder.FullPath);

			foreach (var child in folder.Children.OfType<FolderNode>())
				RemoveThumbStateRecursive(child);
		}

		private void ScheduleLibraryRefreshForPath(string changedPath)
		{
			Dispatcher.Invoke(() =>
			{
				var folderPath = GetFolderPathForChange(changedPath);
				if (!string.IsNullOrWhiteSpace(folderPath))
					_dirtyFolders.Add(folderPath);

				if (_libraryDebounceTimer == null) return;
				_libraryDebounceTimer.Stop();
				_libraryDebounceTimer.Start();
			});
		}

		private string GetFolderPathForChange(string changedPath)
		{
			try
			{
				if (string.IsNullOrWhiteSpace(changedPath))
					return LibraryFolder;

				// If the changed path is a directory, refresh that directory.
				if (Directory.Exists(changedPath))
					return Path.GetFullPath(changedPath);

				// Otherwise refresh the parent folder.
				var dir = Path.GetDirectoryName(changedPath);
				return string.IsNullOrWhiteSpace(dir)
					? LibraryFolder
					: Path.GetFullPath(dir);
			}
			catch
			{
				return LibraryFolder;
			}
		}

		private void RefreshLibraryFromDisk()
		{
			BuildLibraryTreeSkeleton();
			LoadThumbnailsForRootOnly();
			EnsureRootExpanded();
			RestoreExpandedFolders();
		}

        // ---- Build tree skeleton (no subfolder thumbnails) ----

        private void BuildLibraryTreeSkeleton()
        {
            _thumbsLoadedForFolder.Clear();
            LibraryRoot.Clear();

            var root = new FolderNode
            {
                Name = "Library",
                FullPath = LibraryFolder
            };

            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(LibraryFolder, "*.*", SearchOption.AllDirectories)
                                 .Where(f => SupportedExt.Contains(Path.GetExtension(f)));
            }
            catch
            {
                return;
            }

            foreach (var file in files)
                AddFileToTree(root, file);

            SortFolderRecursive(root);
            LibraryRoot.Add(root);
        }

        private void AddFileToTree(FolderNode root, string fullFilePath)
        {
            var relative = Path.GetRelativePath(LibraryFolder, fullFilePath);

            var parts = relative.Split(
                new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0) return;

            var folderParts = parts.Take(parts.Length - 1);
            var filename = parts[^1];

            var current = root;

            foreach (var folderName in folderParts)
            {
                var next = current.Children.OfType<FolderNode>()
                    .FirstOrDefault(f => string.Equals(f.Name, folderName, StringComparison.OrdinalIgnoreCase));

                if (next == null)
                {
                    next = new FolderNode
                    {
                        Name = folderName,
                        FullPath = Path.Combine(current.FullPath ?? LibraryFolder, folderName)
                    };
                    current.Children.Add(next);
                }

                current = next;
            }

            current.Children.Add(new ImageNode
            {
                Name = filename,
                FullPath = fullFilePath,
                Thumbnail = null
            });
        }

        private void SortFolderRecursive(FolderNode folder)
        {
            var folders = folder.Children
                .OfType<FolderNode>()
                .OrderBy(f => f.Name, StringComparer.OrdinalIgnoreCase)
                .Cast<LibraryNode>()
                .ToList();

            var images = folder.Children
                .OfType<ImageNode>()
                .OrderBy(i => i.Name, StringComparer.OrdinalIgnoreCase)
                .Cast<LibraryNode>()
                .ToList();

            folder.Children.Clear();

            foreach (var f in folders) folder.Children.Add(f);
            foreach (var i in images) folder.Children.Add(i);

            foreach (var sub in folder.Children.OfType<FolderNode>())
                SortFolderRecursive(sub);
        }

        // ---- Lazy thumbnail loading ----

		private void LoadThumbnailsForRootOnly()
		{
			if (LibraryRoot.Count == 0) return;
			var root = LibraryRoot[0];

			if (!string.IsNullOrWhiteSpace(root.FullPath))
				_thumbsLoadedForFolder.Add(root.FullPath);

			foreach (var img in root.Children.OfType<ImageNode>())
			{
				if (img.Thumbnail != null) continue;
				_ = LoadThumbnailAsync(img);
			}
		}

		private void LoadThumbnailsForFolder(FolderNode folder)
		{
			if (folder.FullPath == null) return;

			if (_thumbsLoadedForFolder.Contains(folder.FullPath))
				return;

			_thumbsLoadedForFolder.Add(folder.FullPath);

			foreach (var img in folder.Children.OfType<ImageNode>())
			{
				if (img.Thumbnail != null) continue;
				_ = LoadThumbnailAsync(img);
			}
		}

		private async Task LoadThumbnailAsync(ImageNode img)
		{
			var path = img.FullPath;
			var cachePath = GetThumbnailCachePath(path);

			lock (_thumbnailsLoading)
			{
				if (_thumbnailsLoading.Contains(path))
					return;

				_thumbnailsLoading.Add(path);
			}

			try
			{
				await _thumbnailDecodeGate.WaitAsync();

				BitmapSource? thumb = await Task.Run(() =>
				{
					// Try disk cache first
					var cached = LoadBitmapSourceFromFile(cachePath);
					if (cached != null)
						return cached;

					// Otherwise decode original and save cache
					return CreateThumbnailAndPersistCache(path, cachePath);
				});

				await Dispatcher.InvokeAsync(() =>
				{
					if (!string.Equals(img.FullPath, path, StringComparison.OrdinalIgnoreCase))
						return;

					if (img.Thumbnail != null)
						return;

					img.Thumbnail = thumb;
				});
			}
			finally
			{
				_thumbnailDecodeGate.Release();

				lock (_thumbnailsLoading)
				{
					_thumbnailsLoading.Remove(path);
				}
			}
		}

        private static BitmapSource? LoadThumbnail(string path)
        {
            try
            {
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.DecodePixelWidth = 200;
                bi.UriSource = new Uri(path, UriKind.Absolute);
                bi.EndInit();
                bi.Freeze();
                return bi;
            }
            catch
            {
                return null;
            }
        }

        // ---- TreeView events ----

		private void FolderHeader_Click(object sender, MouseButtonEventArgs e)
		{
			if (_folderClickLocked)
			{
				e.Handled = true;
				return;
			}

			_folderClickLocked = true;
			_folderClickTimer.Stop();
			_folderClickTimer.Start();

			if (sender is not DependencyObject dep)
				return;

			var tvi = FindAncestor<TreeViewItem>(dep);
			if (tvi == null || tvi.DataContext is not FolderNode folder || string.IsNullOrWhiteSpace(folder.FullPath))
				return;

			// Root "Library" folder cannot collapse
			if (string.Equals(folder.FullPath, LibraryFolder, StringComparison.OrdinalIgnoreCase))
			{
				tvi.IsExpanded = true;
				e.Handled = true;
				return;
			}

			// If currently expanded, just collapse normally.
			if (tvi.IsExpanded)
			{
				tvi.IsExpanded = false;
				e.Handled = true;
				return;
			}

			// Collapsed -> refresh folder first, then expand refreshed node.
			string folderPath = folder.FullPath;

			RefreshSingleFolder(folderPath);

			Dispatcher.BeginInvoke(new Action(() =>
			{
				if (LibraryRoot.Count == 0)
					return;

				var refreshedFolder = FindFolderNode(LibraryRoot[0], folderPath);
				if (refreshedFolder == null)
					return;

				var refreshedTvi = FindTreeViewItemRecursive(LibraryTree, refreshedFolder);
				if (refreshedTvi != null)
					refreshedTvi.IsExpanded = true;
			}), DispatcherPriority.Loaded);

			e.Handled = true;
		}

		private void LibraryTreeItem_Expanded(object sender, RoutedEventArgs e)
		{
			if (e.OriginalSource is TreeViewItem tvi && tvi.DataContext is FolderNode folder)
			{
				if (!string.IsNullOrWhiteSpace(folder.FullPath))
					_expandedFolders.Add(folder.FullPath);

				LoadThumbnailsForFolder(folder);
			}
		}

		private void LibraryTreeItem_Collapsed(object sender, RoutedEventArgs e)
		{
			if (e.OriginalSource is TreeViewItem tvi && tvi.DataContext is FolderNode folder)
			{
				if (!string.IsNullOrWhiteSpace(folder.FullPath) &&
					!string.Equals(folder.FullPath, LibraryFolder, StringComparison.OrdinalIgnoreCase))
				{
					_expandedFolders.Remove(folder.FullPath);
				}
			}
		}
		
		private void RestoreExpandedFolders()
		{
			if (LibraryRoot.Count == 0)
				return;

			Dispatcher.BeginInvoke(new Action(() =>
			{
				foreach (var root in LibraryRoot)
					RestoreExpandedFoldersRecursive(root);
			}), DispatcherPriority.Loaded);
		}

		private void RestoreExpandedFoldersRecursive(FolderNode folder)
		{
			var tvi = FindTreeViewItem(folder);
			if (tvi != null)
			{
				bool shouldExpand =
					string.Equals(folder.FullPath, LibraryFolder, StringComparison.OrdinalIgnoreCase) ||
					(!string.IsNullOrWhiteSpace(folder.FullPath) && _expandedFolders.Contains(folder.FullPath));

				if (shouldExpand)
					tvi.IsExpanded = true;
			}

			foreach (var child in folder.Children.OfType<FolderNode>())
				RestoreExpandedFoldersRecursive(child);
		}

		private TreeViewItem? FindTreeViewItem(object item)
		{
			return FindTreeViewItemRecursive(LibraryTree, item);
		}

		private TreeViewItem? FindTreeViewItemRecursive(ItemsControl parent, object item)
		{
			if (parent.ItemContainerGenerator.ContainerFromItem(item) is TreeViewItem direct)
				return direct;

			foreach (var childItem in parent.Items)
			{
				if (parent.ItemContainerGenerator.ContainerFromItem(childItem) is TreeViewItem childContainer)
				{
					var result = FindTreeViewItemRecursive(childContainer, item);
					if (result != null)
						return result;
				}
			}

			return null;
		}

        private static T? FindAncestor<T>(DependencyObject start) where T : DependencyObject
        {
            var current = start;
            while (current != null)
            {
                if (current is T match) return match;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

		private void ImageNode_Click(object sender, MouseButtonEventArgs e)
		{
			// Ignore image clicks that happen as part of the same physical click
			// used to expand/collapse a folder header.
			if (_folderClickLocked)
			{
				e.Handled = true;
				return;
			}

			if (sender is FrameworkElement fe && fe.DataContext is ImageNode img)
			{
				OpenLibraryImage(img.FullPath);
				e.Handled = true;
			}
		}

        // ---- Open/close images ----

        private void OpenLibraryImage(string libraryPath)
        {
            if (!File.Exists(libraryPath))
                return;

            if (_openByPath.ContainsKey(libraryPath))
            {
                UpdateHint();
                return;
            }

            BitmapSource bitmap;
            try
            {
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.UriSource = new Uri(libraryPath, UriKind.Absolute);
                bi.EndInit();
                bi.Freeze();
                bitmap = bi;
            }
            catch
            {
                return;
            }

            var item = new ImageItem
            {
                Path = libraryPath,
                Bitmap = bitmap
            };

            Images.Add(item);
            _openByPath[libraryPath] = item;

            FitToScreen();
            UpdateDesignHeight();
            UpdateHint();
        }

        private void CloseOpenImage(ImageItem item)
        {
            Images.Remove(item);
            _openByPath.Remove(item.Path);
            UpdateHint();
        }

        private void ImageItem_RightClickClose(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is ImageItem item)
            {
                CloseOpenImage(item);
                e.Handled = true;
            }
        }
		
		private void ImageItem_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			// Clear any stale suppression left over from the previous drag.
			_suppressNextImageToggleClick = false;

			if (sender is not FrameworkElement fe || fe.DataContext is not ImageItem item)
				return;

			_draggedOpenImage = item;
			_lastDragOverTarget = null;
			_openImageDragStartPoint = e.GetPosition(ImagesItems);
			_openImageDragPending = true;
		}

		private void ImageItem_PreviewMouseMove(object sender, MouseEventArgs e)
		{
			if (!_openImageDragPending || _draggedOpenImage == null)
				return;

			if (e.LeftButton != MouseButtonState.Pressed)
			{
				_openImageDragPending = false;
				_draggedOpenImage = null;
				_lastDragOverTarget = null;
				return;
			}

			var currentPos = e.GetPosition(ImagesItems);

			if (Math.Abs(currentPos.X - _openImageDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
				Math.Abs(currentPos.Y - _openImageDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
			{
				return;
			}

			_openImageDragPending = false;

			var draggedItem = _draggedOpenImage;
			var data = new DataObject(typeof(ImageItem), draggedItem);
			bool dragStarted = false;

			try
			{
				draggedItem.IsBeingDragged = true;
				Mouse.OverrideCursor = Cursors.SizeAll;
				dragStarted = true;

				DragDrop.DoDragDrop((DependencyObject)sender, data, DragDropEffects.Move);
			}
			finally
			{
				draggedItem.IsBeingDragged = false;
				Mouse.OverrideCursor = null;
				_draggedOpenImage = null;
				_lastDragOverTarget = null;

				if (dragStarted)
					_suppressNextImageToggleClick = true;
			}
		}

        private void ImageItem_DragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(ImageItem)))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            if (e.Data.GetData(typeof(ImageItem)) is not ImageItem dragged)
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            if (sender is not FrameworkElement fe || fe.DataContext is not ImageItem target)
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            e.Effects = DragDropEffects.Move;

            if (ReferenceEquals(dragged, target))
            {
                _lastDragOverTarget = target;
                e.Handled = true;
                return;
            }

            if (ReferenceEquals(_lastDragOverTarget, target))
            {
                e.Handled = true;
                return;
            }

            int oldIndex = Images.IndexOf(dragged);
            int newIndex = Images.IndexOf(target);

            if (oldIndex >= 0 && newIndex >= 0 && oldIndex != newIndex)
            {
                Images.Move(oldIndex, newIndex);
            }

            _lastDragOverTarget = target;
            e.Handled = true;
        }

        private void ImageItem_Drop(object sender, DragEventArgs e)
        {
            _lastDragOverTarget = null;
            e.Handled = true;
        }

        private void ImagesItems_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ImageItem)))
            {
                e.Effects = DragDropEffects.Move;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void ImagesItems_Drop(object sender, DragEventArgs e)
        {
            _lastDragOverTarget = null;

            if (!e.Data.GetDataPresent(typeof(ImageItem)))
                return;

            if (e.Data.GetData(typeof(ImageItem)) is not ImageItem dragged)
                return;

            int oldIndex = Images.IndexOf(dragged);
            if (oldIndex < 0)
                return;

            if (oldIndex != Images.Count - 1)
                Images.Move(oldIndex, Images.Count - 1);

            e.Handled = true;
        }

		private void ImageToggle_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
		{
			if (!_suppressNextImageToggleClick)
				return;

			_suppressNextImageToggleClick = false;
			e.Handled = true;
		}
		
		private void StartLibraryPolling()
		{
			//StopLibraryPolling();

			_libraryPollTimer = new DispatcherTimer { Interval = LibraryPollInterval };
			_libraryPollTimer.Tick += (_, __) =>
			{
				_dirtyFolders.Add(LibraryFolder);
				RefreshDirtyFolders();
			};
			_libraryPollTimer.Start();
		}

		private void StopLibraryPolling()
		{
			if (_libraryPollTimer != null)
			{
				_libraryPollTimer.Stop();
				_libraryPollTimer = null;
			}
		}
    }
}