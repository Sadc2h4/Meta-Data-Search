# Meta-Data-Search
![type](https://img.shields.io/badge/app-Windows%20Forms-blue?style=flat-square&logo=windows)
![runtime](https://img.shields.io/badge/.NET-6.0%20(or%208.0)-512BD4?style=flat-square&logo=dotnet)
![tfm](https://img.shields.io/badge/TFM-net6.0--windows-0B7285?style=flat-square)

<img width="720" alt="GitHub_0" src="https://github.com/user-attachments/assets/14b6d160-1791-4be8-9ebd-cddfdc42412e" />

## Download
<a href="https://github.com/Sadc2h4/Meta-Data-Search/releases/tag/V1.1a">
  <img
    src="https://raw.githubusercontent.com/Sadc2h4/brand-assets/main/button/Download_Button_1.png"
    alt="Download .zip"
    height="48"
  />
</a>

<br>

## Features
This application was created to extract generation information from the metadata within output images,  
as it was cumbersome to manually store all the details—such as the prompts and Lora weights used during  
generation when creating images with Stable Diffusion-based generative AI.

It extracts metadata from Stable Diffusion images (ComfyUI / AUTOMATIC1111, etc.) and  
supports loading PNG / JPG / WebP files (with preview support) via drag-and-drop or buttons.

This is the light version, which omits the AI prompt functionality to reduce the overall file size.  
If you primarily want to use it for metadata extraction, this version might be preferable.

## Feasible functions

https://github.com/user-attachments/assets/41df9f8c-6f60-445f-8daf-a608fabf6c21

・Extract and format metadata within images  
　Checkpoint / LoRA (Name・ID/URN・Weight)  
　Prompt / Negative Prompt  
　Steps / Sampler / CFG / Denoise / Size / Scale / Seed  
　Embedding (e.g., EasyNegative)  

・Model Hash / LoRA Hash for models/LoRA (when available)  

・Switch between grid (tabular) display and text display  

・Copy values to clipboard using the “copy” button on the right side of each value  
　Checkpoint / LoRA / Embedding values link to the corresponding page on Civitai  

・Fallback to Civitai search when ID/URN cannot be obtained directly  

・Raw data dump display (integrated display of XMP / PNG text / EXIF)  

・Automatically resized preview (including WebP) in the upper right corner  

・For images without metadata, generate pseudo prompts with **AI Caption (offline)** (optional)  

## How to Use
1. Launch Meta_DataSearch.exe

2. Select an image using “Select image” (drag-and-drop also works)

3. The preview appears in the top-right and metadata is displayed at the bottom

4. “Table mode”: Table format (copy/link)

5. “Display raw data”: Show raw metadata

6. Save JSON using “Save raw data” if needed

## Deletion Method
・Please delete the entire file.

## Disclaimer
・I assume no responsibility whatsoever for any damages incurred through the use of this file.
