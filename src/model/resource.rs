//! Resource (image, media) model structures.

use serde::{Deserialize, Serialize};

/// Type of resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum ResourceType {
    /// Image (PNG, JPEG, GIF, BMP, TIFF, WMF, EMF)
    Image,
    /// Audio file
    Audio,
    /// Video file
    Video,
    /// Chart (extracted as image)
    Chart,
    /// Embedded OLE object
    Ole,
    /// Other binary data
    Other,
}

impl ResourceType {
    /// Determine resource type from MIME type.
    pub fn from_mime_type(mime: &str) -> Self {
        let mime_lower = mime.to_lowercase();
        if mime_lower.starts_with("image/") {
            ResourceType::Image
        } else if mime_lower.starts_with("audio/") {
            ResourceType::Audio
        } else if mime_lower.starts_with("video/") {
            ResourceType::Video
        } else if mime_lower.contains("chart") {
            ResourceType::Chart
        } else if mime_lower.contains("ole") || mime_lower.contains("oleobject") {
            ResourceType::Ole
        } else {
            ResourceType::Other
        }
    }

    /// Determine resource type from file extension.
    pub fn from_extension(ext: &str) -> Self {
        match ext.to_lowercase().as_str() {
            "png" | "jpg" | "jpeg" | "gif" | "bmp" | "tiff" | "tif" | "wmf" | "emf" | "svg" => {
                ResourceType::Image
            }
            "mp3" | "wav" | "ogg" | "m4a" | "wma" => ResourceType::Audio,
            "mp4" | "avi" | "mov" | "wmv" | "webm" => ResourceType::Video,
            _ => ResourceType::Other,
        }
    }
}

/// A binary resource (image, media file, etc.).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Resource {
    /// Resource type
    pub resource_type: ResourceType,

    /// Original filename (if known)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filename: Option<String>,

    /// MIME type
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,

    /// Binary data
    #[serde(skip)]
    pub data: Vec<u8>,

    /// Size in bytes
    pub size: usize,

    /// Width in pixels (for images)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<u32>,

    /// Height in pixels (for images)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<u32>,

    /// Alt text / description
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,
}

impl Resource {
    /// Create a new resource.
    pub fn new(resource_type: ResourceType, data: Vec<u8>) -> Self {
        let size = data.len();
        Self {
            resource_type,
            filename: None,
            mime_type: None,
            data,
            size,
            width: None,
            height: None,
            alt_text: None,
        }
    }

    /// Create an image resource.
    pub fn image(data: Vec<u8>, filename: Option<String>) -> Self {
        let size = data.len();
        let mime_type = filename
            .as_ref()
            .and_then(|f| Self::mime_from_filename(f));
        Self {
            resource_type: ResourceType::Image,
            filename,
            mime_type,
            data,
            size,
            width: None,
            height: None,
            alt_text: None,
        }
    }

    /// Get the file extension for this resource.
    pub fn extension(&self) -> Option<&str> {
        self.filename.as_ref().and_then(|f| {
            f.rsplit('.')
                .next()
                .filter(|ext| ext.len() <= 5 && ext.chars().all(|c| c.is_alphanumeric()))
        })
    }

    /// Determine MIME type from filename.
    pub fn mime_from_filename(filename: &str) -> Option<String> {
        let ext = filename.rsplit('.').next()?.to_lowercase();
        let mime = match ext.as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" | "tif" => "image/tiff",
            "svg" => "image/svg+xml",
            "wmf" => "image/x-wmf",
            "emf" => "image/x-emf",
            "mp3" => "audio/mpeg",
            "wav" => "audio/wav",
            "ogg" => "audio/ogg",
            "m4a" => "audio/mp4",
            "mp4" => "video/mp4",
            "avi" => "video/x-msvideo",
            "mov" => "video/quicktime",
            "webm" => "video/webm",
            _ => return None,
        };
        Some(mime.to_string())
    }

    /// Generate a suggested filename for this resource.
    pub fn suggested_filename(&self, id: &str) -> String {
        if let Some(ref filename) = self.filename {
            filename.clone()
        } else {
            let ext = match self.resource_type {
                ResourceType::Image => self
                    .mime_type
                    .as_ref()
                    .and_then(|m| Self::extension_from_mime(m))
                    .unwrap_or("png"),
                ResourceType::Audio => "mp3",
                ResourceType::Video => "mp4",
                ResourceType::Chart => "png",
                ResourceType::Ole => "bin",
                ResourceType::Other => "bin",
            };
            format!("{}.{}", id, ext)
        }
    }

    /// Get extension from MIME type.
    fn extension_from_mime(mime: &str) -> Option<&'static str> {
        match mime {
            "image/png" => Some("png"),
            "image/jpeg" => Some("jpg"),
            "image/gif" => Some("gif"),
            "image/bmp" => Some("bmp"),
            "image/tiff" => Some("tiff"),
            "image/svg+xml" => Some("svg"),
            "image/x-wmf" => Some("wmf"),
            "image/x-emf" => Some("emf"),
            "audio/mpeg" => Some("mp3"),
            "audio/wav" => Some("wav"),
            "video/mp4" => Some("mp4"),
            _ => None,
        }
    }

    /// Save resource to a file.
    pub fn save_to(&self, path: impl AsRef<std::path::Path>) -> std::io::Result<()> {
        std::fs::write(path, &self.data)
    }

    /// Check if this is an image.
    pub fn is_image(&self) -> bool {
        matches!(self.resource_type, ResourceType::Image | ResourceType::Chart)
    }

    /// Check if this is a media file (audio/video).
    pub fn is_media(&self) -> bool {
        matches!(
            self.resource_type,
            ResourceType::Audio | ResourceType::Video
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resource_type_from_mime() {
        assert_eq!(
            ResourceType::from_mime_type("image/png"),
            ResourceType::Image
        );
        assert_eq!(
            ResourceType::from_mime_type("IMAGE/JPEG"),
            ResourceType::Image
        );
        assert_eq!(
            ResourceType::from_mime_type("audio/mpeg"),
            ResourceType::Audio
        );
        assert_eq!(
            ResourceType::from_mime_type("video/mp4"),
            ResourceType::Video
        );
        assert_eq!(
            ResourceType::from_mime_type("application/octet-stream"),
            ResourceType::Other
        );
    }

    #[test]
    fn test_resource_type_from_extension() {
        assert_eq!(ResourceType::from_extension("png"), ResourceType::Image);
        assert_eq!(ResourceType::from_extension("JPG"), ResourceType::Image);
        assert_eq!(ResourceType::from_extension("mp3"), ResourceType::Audio);
        assert_eq!(ResourceType::from_extension("mp4"), ResourceType::Video);
        assert_eq!(ResourceType::from_extension("xyz"), ResourceType::Other);
    }

    #[test]
    fn test_resource_creation() {
        let data = vec![0x89, 0x50, 0x4E, 0x47]; // PNG magic
        let resource = Resource::image(data.clone(), Some("test.png".to_string()));

        assert_eq!(resource.resource_type, ResourceType::Image);
        assert_eq!(resource.size, 4);
        assert_eq!(resource.filename, Some("test.png".to_string()));
        assert_eq!(resource.mime_type, Some("image/png".to_string()));
    }

    #[test]
    fn test_resource_extension() {
        let resource = Resource::image(vec![], Some("image.png".to_string()));
        assert_eq!(resource.extension(), Some("png"));

        let resource2 = Resource::image(vec![], Some("photo.JPEG".to_string()));
        assert_eq!(resource2.extension(), Some("JPEG"));
    }

    #[test]
    fn test_suggested_filename() {
        let resource = Resource::image(vec![], Some("original.png".to_string()));
        assert_eq!(resource.suggested_filename("img1"), "original.png");

        let mut resource2 = Resource::new(ResourceType::Image, vec![]);
        resource2.mime_type = Some("image/jpeg".to_string());
        assert_eq!(resource2.suggested_filename("img2"), "img2.jpg");
    }

    #[test]
    fn test_is_image() {
        let image = Resource::new(ResourceType::Image, vec![]);
        assert!(image.is_image());

        let chart = Resource::new(ResourceType::Chart, vec![]);
        assert!(chart.is_image());

        let audio = Resource::new(ResourceType::Audio, vec![]);
        assert!(!audio.is_image());
    }

    #[test]
    fn test_is_media() {
        let audio = Resource::new(ResourceType::Audio, vec![]);
        assert!(audio.is_media());

        let video = Resource::new(ResourceType::Video, vec![]);
        assert!(video.is_media());

        let image = Resource::new(ResourceType::Image, vec![]);
        assert!(!image.is_media());
    }
}
