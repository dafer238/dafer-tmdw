# Build and Release Workflow

This document explains how to build the project locally and create releases on GitHub.

## Project Structure

```
dafer-tmdw/
├── src/                      # Source code
│   ├── CoolPropWrapper.cs    # Main wrapper code
│   ├── CoolPropWrapper.csproj # Project file
│   ├── CoolPropWrapper.dna   # ExcelDNA configuration
│   ├── CoolProp.dll          # CoolProp library dependency
│   ├── bin/                  # Build output (ignored by git)
│   └── obj/                  # Build artifacts (ignored by git)
├── compiled/                 # Release-ready files (tracked in git)
│   └── net48/               # .NET Framework 4.8 build
│       ├── CoolProp.dll
│       ├── CoolPropWrapper.dll
│       ├── CoolPropWrapper.dna
│       └── CoolPropWrapper.xll
├── .github/workflows/        # GitHub Actions workflows
│   └── release.yml          # Automated release workflow
├── build.bat                # Build script
├── README.md                # Main documentation
└── RELEASE.md               # This file
```

## Building Locally

### Prerequisites

- .NET SDK installed (check with `dotnet --version`)
- Windows operating system

### Build Steps

1. Run the build script:

   ```cmd
   build.bat
   ```

2. The script will:
   - Clean previous builds
   - Restore NuGet packages
   - Build the project for .NET Framework 4.8 (x64)
   - Copy all necessary files to `compiled/net48/`

3. Verify the output in `compiled/net48/`:
   - `CoolProp.dll`
   - `CoolPropWrapper.dll`
   - `CoolPropWrapper.dna`
   - `CoolPropWrapper.xll`

## Creating a Release

### Step 1: Build and Commit

After running `build.bat`, commit the built files:

```cmd
git add compiled/net48/*
git commit -m "Build version X.X.X"
```

### Step 2: Create and Push a Version Tag

Create a tag following semantic versioning (v0.1.1, v0.2.0, etc.):

```cmd
git tag v0.1.1
git push origin v0.1.1
```

### Step 3: Automatic Release

Once you push the tag, GitHub Actions will automatically:

1. Detect the version tag
2. Package the files from `compiled/net48/` into a ZIP archive
3. Create a GitHub release with the version number
4. Attach the ZIP file as a release asset

### Example Workflow

```cmd
# 1. Build the project
build.bat

# 2. Commit the changes
git add compiled/net48/*
git commit -m "Build version 0.1.2"

# 3. Create and push tag
git tag v0.1.2
git push origin v0.1.2
```

Within a few minutes, a new release will appear at:
`https://github.com/dafer238/dafer-tmdw/releases/tag/v0.1.2`

## Manual Release Trigger

You can also trigger a release manually from the GitHub Actions tab:

1. Go to the repository on GitHub
2. Click on "Actions"
3. Select "Create Release" workflow
4. Click "Run workflow"
5. Enter the tag name (e.g., `v0.1.3`) in the input field
6. Click "Run workflow" to confirm

**Note:** The tag you specify will be created if it doesn't already exist. Make sure you've committed your built files to `compiled/net48/` before triggering the manual release.

## Version Numbering

Follow semantic versioning:

- **v0.1.0** - Initial release
- **v0.1.1** - Patch (bug fixes)
- **v0.2.0** - Minor (new features, backwards compatible)
- **v1.0.0** - Major (breaking changes)

## Troubleshooting

### Build fails

- Ensure .NET SDK is installed
- Run `dotnet restore` manually to check for package issues

### Release not created

- Verify the tag follows the format `vX.Y.Z` (e.g., `v0.1.1`)
- Check GitHub Actions tab for workflow errors
- Ensure `compiled/net48/` contains all required files

### Files missing in compiled folder

- The `compiled/` folder is tracked in git (not in `.gitignore`)
- Commit and push changes in `compiled/net48/` before tagging
